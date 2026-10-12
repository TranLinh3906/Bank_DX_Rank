import unittest
import tempfile
from pathlib import Path
from unittest.mock import patch

import pandas as pd

import scoring


class CandidateScoringTests(unittest.TestCase):
    def evaluate(self, criterion, context, is_tier2=False):
        row = {
            "text_normalized": scoring.normalize_text(context),
            "is_tier2": is_tier2
        }
        return scoring.evaluate_candidate(row, criterion)

    def test_cooccurring_topic_and_generic_process_do_not_score_two(self):
        score = self.evaluate(
            "D1",
            "RPA duoc su dung de tu dong hoa quy trinh xu ly ho so.",
            is_tier2=True
        )[0]
        self.assertEqual(score, 1)

    def test_explicit_multiple_processes_can_score_two(self):
        score = self.evaluate(
            "D1",
            "RPA tu dong hoa nhieu quy trinh nghiep vu toan hang."
        )[0]
        self.assertEqual(score, 2)

    def test_tier2_filter_stage_does_not_change_score(self):
        context = "RPA duoc su dung de tu dong hoa quy trinh xu ly ho so."
        regular = self.evaluate("D1", context, is_tier2=False)[0]
        tier2 = self.evaluate("D1", context, is_tier2=True)[0]
        self.assertEqual(regular, tier2)

    def test_negated_evidence_is_not_positive_support(self):
        score = self.evaluate(
            "A2",
            "Ngan hang chua co dau tu vao ha tang cong nghe."
        )[0]
        self.assertEqual(score, 0)

    def test_upstream_github_evidence_is_preferred_when_present(self):
        with tempfile.TemporaryDirectory() as directory:
            project = Path(directory) / "BANK_DX_PROJECT"
            upstream = (
                Path(directory)
                / "Digital-banking-scoring-system"
                / "Evidence_Final.xlsx"
            )
            upstream.parent.mkdir(parents=True)
            upstream.touch()

            selected = scoring.resolve_evidence_path(project)

            self.assertEqual(selected, upstream)

    def test_local_evidence_is_fallback_without_upstream_repo(self):
        with tempfile.TemporaryDirectory() as directory:
            project = Path(directory) / "BANK_DX_PROJECT"
            local = project / "data" / "evidence.xlsx"
            local.parent.mkdir(parents=True)
            local.touch()

            selected = scoring.resolve_evidence_path(project)

            self.assertEqual(selected, local)

    def test_money_amount_without_currency_is_not_investment_scale(self):
        score = self.evaluate(
            "A2",
            "Dau tu cong nghe giup kich hoat 1,6 trieu vi dien tu."
        )[0]
        self.assertEqual(score, 1)

    def test_generic_strategy_does_not_match_digital_strategy_criterion(self):
        score = self.evaluate(
            "A1",
            "Chien luoc hop tac voi IFC huong den muc tieu phat trien ben vung."
        )[0]
        self.assertEqual(score, 0)

    def test_generic_security_standard_is_not_level_two(self):
        score = self.evaluate(
            "D3",
            "Ngan hang dap ung tieu chuan quoc te ve an toan thong tin."
        )[0]
        self.assertEqual(score, 1)


class RankingTests(unittest.TestCase):
    def test_competition_ranking_preserves_ties(self):
        totals = [32, 32, 24, 24, 16]
        rows = []
        for bank, total in zip(scoring.BANKS, totals):
            remaining = total
            for criterion in scoring.CODES:
                score = min(2, remaining)
                rows.append({
                    "bank": bank,
                    "criterion": criterion,
                    "approved_score": score
                })
                remaining -= score

        _, _, _, ranking = scoring.calculate_scores(
            pd.DataFrame(rows), "approved_score"
        )
        self.assertEqual(
            ranking["rank"].tolist(),
            [1, 1, 3, 3, 5]
        )
        self.assertTrue(ranking["status"].eq("automated").all())
        self.assertTrue(
            ranking["methodology"]
            .eq(scoring.METHODOLOGY_LABEL)
            .all()
        )
        self.assertIn("warning_count", ranking.columns)
        self.assertIn("low_confidence_count", ranking.columns)

    def test_review_output_uses_suggested_scores_and_auto_approves(self):
        review = pd.DataFrame({
            "bank": ["BIDV"],
            "criterion": ["A1"],
            "suggested_score": [1],
            "status": ["automated"],
            "approved_score": [2],
            "reviewer": ["old reviewer"],
            "review_note": ["old note"]
        })
        with tempfile.TemporaryDirectory() as directory:
            review_path = Path(directory) / "review_scores.xlsx"
            with patch.object(scoring, "REVIEW_PATH", review_path):
                scoring.save_review(review)

            saved = pd.read_excel(review_path)

        self.assertEqual(saved.loc[0, "status"], "automated")
        self.assertNotIn("approved_score", saved.columns)
        self.assertNotIn("reviewer", saved.columns)
        self.assertNotIn("review_note", saved.columns)


if __name__ == "__main__":
    unittest.main()
