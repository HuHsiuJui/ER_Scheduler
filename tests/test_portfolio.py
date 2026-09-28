import unittest
import json
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import Mock, patch
from dataclasses import replace
from datetime import date
from leave_accounting import LeaveInput, settle, cycle_for
from main import synthetic_input, parse_input, run
from workbook_export import safe
import scheduler_core as core


class LeaveTests(unittest.TestCase):
    def setUp(self):
        self.item=LeaveInput('DEMO',80,5.07,additional_leave_hours=8,
            cycle_verified=True,regular_rest_remaining=0)

    def test_no_consent_no_deduction(self):
        r=settle(self.item)
        self.assertEqual((r['annual_projected'],r['comp_projected']),(80,5.07))
        self.assertEqual(r['annual_shortfall_hours'],2.93)

    def test_annual_must_be_asked(self):
        r=settle(replace(self.item,comp_consent='同意',annual_consent='同意'))
        self.assertEqual((r['annual_projected'],r['comp_projected']),(80,0))

    def test_both_consents(self):
        r=settle(replace(self.item,comp_consent='同意',annual_consent='同意',annual_asked=True))
        self.assertEqual((r['annual_projected'],r['comp_projected']),(77.07,0))

    def test_pay_does_not_credit(self):
        self.assertEqual(settle(replace(self.item,overtime_hours=8,overtime_choice='領加班費'))['comp_projected'],5.07)

    def test_future_credit_cannot_be_spent(self):
        r=settle(replace(self.item,overtime_hours=8,overtime_choice='轉補休'))
        self.assertEqual(r['comp_projected'],13.07)
        self.assertEqual(r['suggested_comp_hours'],5.07)

    def test_ordinary_rest_first(self):
        r=settle(replace(self.item,regular_rest_remaining=1,comp_consent='同意'))
        self.assertIsNone(r['suggested_comp_hours'])
        self.assertEqual(r['comp_projected'],5.07)

    def test_missing_not_zero(self):
        self.assertIsNone(settle(replace(self.item,comp_opening=None))['comp_projected'])

    def test_annual_insufficient(self):
        r=settle(replace(self.item,annual_opening=0,comp_consent='同意',annual_asked=True,annual_consent='同意'))
        self.assertEqual(r['annual_projected'],0)
        self.assertIsNone(r['annual_official'])

    def test_refuse_comp(self):
        r=settle(replace(self.item,comp_consent='不同意'))
        self.assertEqual(r['annual_shortfall_hours'],8)

    def test_negative_rejected(self):
        with self.assertRaises(ValueError): settle(replace(self.item,comp_opening=-1))

    def test_cycle_anchor(self):
        self.assertEqual(cycle_for(date(2026,9,1)),(date(2026,8,13),date(2026,9,9)))
        self.assertEqual(cycle_for(date(2026,9,10)),(date(2026,9,10),date(2026,10,7)))


class InputTests(unittest.TestCase):
    def test_demo_complete(self):
        nurses,reqs,boundary,training,config=parse_input(synthetic_input())
        self.assertEqual(len(nurses),38)
        self.assertEqual(len(boundary),38*14)
        self.assertEqual(len(config.preceptor_ids),38)

    def test_missing_boundary_rejected(self):
        d=synthetic_input(); d['boundary'].pop()
        with self.assertRaises(ValueError): parse_input(d)

    def test_duplicate_request_rejected(self):
        d=synthetic_input(); d['requests'].append(d['requests'][0])
        with self.assertRaises(ValueError): parse_input(d)

    def test_excel_injection(self):
        self.assertEqual(safe('=1+1'),"'=1+1")

    def test_administrative_shift_is_attendance(self):
        self.assertTrue(core.is_attendance_shift('8-5'))

    def test_observation_labels(self):
        self.assertEqual(core.canonical_area('Observation'),'Observation')


class EngineRegressionTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.data=synthetic_input()
        cls.result,cls.requests,cls.report=run(cls.data)
        cls.nurses,_,cls.boundary,cls.training,cls.config=parse_input(cls.data)

    def codes(self,result=None,boundary=None):
        check=core.validate_result(result or self.result,self.requests,
            self.boundary if boundary is None else boundary,self.config)
        return {i.code for i in check.hard_errors}

    def test_feasible_smoke(self):
        self.assertEqual(self.result.solver_status,'OPTIMAL')
        self.assertEqual(len(self.result.assignments),570)
        self.assertEqual(self.report['schedule_constraint_errors'],[])

    def test_local_never_claims_clinical_pass(self):
        self.assertFalse(self.report['clinical_publishable'])
        self.assertIn('GOOGLE_SOURCE',self.codes())
        self.assertIn('PREVIOUS_GOOGLE_SOURCE',self.codes())

    def test_zero_cross_and_overtime_for_demo_only(self):
        self.assertEqual(self.result.solver_metrics['overtime_shifts'],0)
        self.assertEqual(self.result.solver_metrics['max_cross_shift_days_per_person_actual'],0)

    def test_cross_month_six_days_caught(self):
        boundary=tuple(replace(b,shift='D',raw='行政或白班',attendance=True)
            if b.nurse_id=='D001' and b.day in {date(2026,8,30),date(2026,8,31)} else b
            for b in self.boundary)
        result=replace(self.result,assignments=self.result.assignments+(
            core.Assignment('D001',date(2026,9,4),'D','Clinic3'),))
        self.assertIn('MAX_CONSECUTIVE',self.codes(result,boundary))

    def test_hard_off_violation_caught(self):
        result=replace(self.result,assignments=self.result.assignments+(
            core.Assignment('D001',date(2026,9,4),'D','Clinic3'),))
        self.assertIn('HARD_OFF_VIOLATION',self.codes(result))

    def test_core_missing_caught(self):
        result=replace(self.result,assignments=self.result.assignments[1:])
        self.assertIn('CORE_ROLE_SHORTAGE',self.codes(result))

    def test_teacher_missing_caught(self):
        changed=replace(self.result.assignments[0],formal_teaching=True,preceptor_id=None)
        result=replace(self.result,assignments=(changed,)+self.result.assignments[1:])
        self.assertIn('TEACHING_NO_PRECEPTOR',self.codes(result))

    def test_part_time_target_not_flat_ten(self):
        nurse=replace(self.nurses[0],part_time=True)
        targets=core.build_target_shifts((nurse,),self.requests,self.training,year=2026,month=9,
            official_holidays=frozenset(),part_time_min_shifts=10)
        self.assertEqual(targets[nurse.nurse_id],15)


class IntegrationDiagnosticTests(unittest.TestCase):
    def test_live_open_and_permission_diagnostic_offline(self):
        with TemporaryDirectory() as directory:
            credentials = Path(directory) / 'service_account.json'
            credentials.write_text(json.dumps({'client_email': 'demo@example.invalid'}), encoding='utf-8')
            client = Mock()
            with patch.object(core, '_google_readonly_client', return_value=client):
                sheet, sid = core._open_live_google_spreadsheet('demo-source', credentials, source_label='測試來源')
                self.assertIs(sheet, client.open_by_key.return_value)
                self.assertEqual(sid, 'demo-source')
                client.open_by_key.side_effect = PermissionError('denied')
                with self.assertRaisesRegex(PermissionError, 'demo@example.invalid'):
                    core._open_live_google_spreadsheet('demo-source', credentials, source_label='測試來源')

    def test_disabled_live_source_fails_before_authentication(self):
        with patch.object(core, '_google_readonly_client') as client:
            with self.assertRaisesRegex(ValueError, '未啟用'):
                core._open_live_google_spreadsheet('', Path('unused.json'), source_label='測試來源')
            client.assert_not_called()

    def test_missing_account_email_is_a_controlled_error(self):
        with TemporaryDirectory() as directory:
            credentials = Path(directory) / 'service_account.json'
            credentials.write_text('{}', encoding='utf-8')
            with self.assertRaisesRegex(ValueError, 'client_email'):
                core._service_account_email(credentials)

    def test_timeout_advice_does_not_claim_infeasibility(self):
        advice = core.build_schedule_advice(solver_error='MILP不可行/未取得可行解: Time limit reached')
        self.assertEqual(advice[0]['category'], '求解未取得可行解')
        self.assertIn('逾時不等於已證明無解', advice[0]['recommendation'])

    def test_training_source_does_not_claim_an_unread_workbook(self):
        nurse = replace(parse_input(synthetic_input())[0][0], hire_date=date(2026,8,1), training_completed_seed=18)
        state = core.load_training_state_canonical((nurse,), year=2026, month=9, interactive=False)[nurse.nurse_id]
        self.assertIn('training_completed_seed', state.source)
        self.assertNotIn('Nurses.xlsx', state.source)


if __name__=='__main__': unittest.main()
