import unittest
import uuid
from pathlib import Path
from unittest.mock import patch

import pandas as pd

import core


class CoreLogicTests(unittest.TestCase):
    def _workspace_tempdir(self):
        tmp_root = Path.cwd() / "tests_tmp"
        tmp_root.mkdir(exist_ok=True)
        tmp_path = tmp_root / f"run_{uuid.uuid4().hex}"
        tmp_path.mkdir(parents=True, exist_ok=True)
        return tmp_path

    def test_parse_oth_col_name_supports_multiple_patterns(self):
        self.assertEqual(core._parse_oth_col_name("s14_1_93_oth"), ("s14_1", 93))
        self.assertEqual(core._parse_oth_col_name("q0_oth95"), ("q0", 95))
        self.assertEqual(core._parse_oth_col_name("Q7Oth12"), ("Q7", 12))
        self.assertEqual(core._parse_oth_col_name("f3_oth"), ("f3", None))

    def test_find_oth_pairs_is_case_insensitive(self):
        df = pd.DataFrame(columns=["Q1", "Q1_Oth", "Q2", "Q2_OTH", "Q3"])
        self.assertEqual(
            core.find_oth_pairs(df),
            [("Q1", "Q1_Oth"), ("Q2", "Q2_OTH")],
        )

    def test_build_spss_column_mapping_supports_order_fallback(self):
        variable_labels = {
            "l_1_s10": "SPSS label s10",
            "l_1_s11": "SPSS label s11",
        }
        value_labels = {
            "l_1_s10": {97: "อื่นๆ ระบุ"},
            "l_1_s11": {98: "อื่นๆ ระบุ"},
        }

        no_fallback = core._build_spss_column_mapping(
            ["เนสกาแฟ_s10", "เนสกาแฟ_s11"],
            variable_labels,
            value_labels,
            allow_order_fallback=False,
        )
        with_fallback = core._build_spss_column_mapping(
            ["เนสกาแฟ_s10", "เนสกาแฟ_s11"],
            variable_labels,
            value_labels,
            allow_order_fallback=True,
        )

        self.assertEqual(no_fallback.unresolved_excel_cols, [])
        self.assertEqual(
            no_fallback.direct_matches,
            [("เนสกาแฟ_s10", "l_1_s10"), ("เนสกาแฟ_s11", "l_1_s11")],
        )
        self.assertEqual(
            with_fallback.direct_matches,
            [("เนสกาแฟ_s10", "l_1_s10"), ("เนสกาแฟ_s11", "l_1_s11")],
        )
        self.assertEqual(with_fallback.fallback_matches, [])
        self.assertEqual(
            with_fallback.resolved_value_labels["เนสกาแฟ_s10"],
            {97: "อื่นๆ ระบุ"},
        )

    def test_build_spss_column_mapping_matches_tail_token_like_new_s2_to_s2(self):
        variable_labels = {"S2": "Question S2"}
        value_labels = {"S2": {97: "อื่นๆ ระบุ"}}

        report = core._build_spss_column_mapping(
            ["New_s2"],
            variable_labels,
            value_labels,
            allow_order_fallback=False,
        )

        self.assertEqual(report.unresolved_excel_cols, [])
        self.assertEqual(report.direct_matches, [("New_s2", "S2")])
        self.assertEqual(report.mismatched_matches, [("New_s2", "S2")])
        self.assertEqual(report.resolved_var_labels["New_s2"], "Question S2")

    def test_build_coding_df_maps_indexed_oth_by_exact_other_code(self):
        df = pd.DataFrame(
            [
                {
                    "Sbjnum": "230066030",
                    "s14_1_O2": "93",
                    "s14_1_O4": "94",
                    "s14_1_O6": "",
                    "s14_1_O8": "",
                    "s14_1_O10": "",
                    "s14_1_93_oth": "fhm hg",
                    "s14_1_94_oth": "fgn",
                }
            ]
        )
        variable_labels = {
            "s14_1_O2": "Question 2",
            "s14_1_O4": "Question 4",
            "s14_1_O6": "Question 6",
            "s14_1_O8": "Question 8",
            "s14_1_O10": "Question 10",
        }
        other_vl = {
            93: "อื่นๆ ระบุ",
            94: "อื่นๆ ระบุ",
        }
        value_labels = {col: other_vl for col in variable_labels}

        coding_df, detected = core.build_coding_df(df, [], variable_labels, value_labels)
        result = coding_df[["Question", "Other_Code", "Open_Text", "Open_Text_From"]]

        self.assertEqual(
            set(map(tuple, result.itertuples(index=False, name=None))),
            {
                ("s14_1_O2", "93", "fhm hg", "s14_1_93_oth"),
                ("s14_1_O4", "94", "fgn", "s14_1_94_oth"),
            },
        )
        self.assertEqual(detected["s14_1_O2"], ["93", "94"])
        self.assertEqual(detected["s14_1_O4"], ["93", "94"])

    def test_build_coding_df_keeps_verbatim_only_rows_and_prefills_cut(self):
        df = pd.DataFrame(
            [
                {
                    "Sbjnum": "9999",
                    "s6": "",
                    "s6_oth": "ตอบมาแต่ไม่ได้ติ๊กโค้ด",
                }
            ]
        )
        variable_labels = {"s6": "Question 6"}
        value_labels = {"s6": {97: "อื่นๆ ระบุ"}}

        coding_df, _ = core.build_coding_df(df, [], variable_labels, value_labels)

        self.assertEqual(len(coding_df), 1)
        row = coding_df.iloc[0]
        self.assertEqual(row["Question"], "s6")
        self.assertEqual(row["Other_Code"], "")
        self.assertEqual(row["Open_Text"], "ตอบมาแต่ไม่ได้ติ๊กโค้ด")
        self.assertEqual(row[core.NEW_OPEN_TEXT_COL], "ตัด")

    def test_build_coding_df_keeps_indexed_verbatim_only_rows_and_prefills_cut(self):
        df = pd.DataFrame(
            [
                {
                    "Sbjnum": "9998",
                    "q0_95": "",
                    "q0_95_oth": "verbatim only",
                }
            ]
        )
        variable_labels = {"q0_95": "Question 0_95"}
        value_labels = {"q0_95": {95: "อื่นๆ ระบุ"}}

        coding_df, _ = core.build_coding_df(
            df,
            core.find_oth_pairs(df),
            variable_labels,
            value_labels,
        )

        self.assertEqual(len(coding_df), 1)
        row = coding_df.iloc[0]
        self.assertEqual(row["Question"], "q0_95")
        self.assertEqual(row["Other_Code"], "")
        self.assertEqual(row["Open_Text"], "verbatim only")
        self.assertEqual(row[core.NEW_OPEN_TEXT_COL], "ตัด")

    def test_merge_coding_keeps_existing_rows_first_and_appends_new_rows(self):
        existing_df = pd.DataFrame(
            [
                {
                    "Question": "s6",
                    core.SBJNUM_COL: "231038341",
                    "Other_Code": "97",
                    "Open_Text": "xxx",
                    "Open_Text_From": "s6_oth",
                    core.NEW_CODE_COL: "500",
                }
            ]
        )
        new_df = pd.DataFrame(
            [
                {
                    "Question": "s6",
                    core.SBJNUM_COL: "9999",
                    "Other_Code": "97",
                    "Open_Text": "test ระบุ",
                    "Open_Text_From": "s6_oth",
                    core.NEW_CODE_COL: "",
                }
            ]
        )

        merged = core._merge_coding_with_existing(existing_df, new_df)
        sorted_df = core._sort_coding_rows(merged)

        self.assertEqual(sorted_df.iloc[0][core.SBJNUM_COL], "231038341")
        self.assertEqual(sorted_df.iloc[0][core.NEW_CODE_COL], "500")
        self.assertEqual(sorted_df.iloc[1][core.SBJNUM_COL], "9999")

    def test_merge_coding_dedupes_blank_values_from_existing_sheet(self):
        existing_df = pd.DataFrame(
            [
                {
                    "Question": "s11_3_O4",
                    "Variable_Label": "Question s11_3",
                    core.SBJNUM_COL: "231161173",
                    "Other_Label": "อื่นๆ ระบุ",
                    "Other_Code": "9826",
                    core.NEW_CODE_COL: "",
                    "Open_Text": None,
                    core.NEW_OPEN_TEXT_COL: None,
                    "Open_Text_From": None,
                    "Remark": None,
                }
            ]
        )
        new_df = pd.DataFrame(
            [
                {
                    "Question": "s11_3_O4",
                    "Variable_Label": "Question s11_3",
                    core.SBJNUM_COL: "231161173",
                    "Other_Label": "อื่นๆ ระบุ",
                    "Other_Code": "9826",
                    core.NEW_CODE_COL: "",
                    "Open_Text": "",
                    core.NEW_OPEN_TEXT_COL: "",
                    "Open_Text_From": "",
                    "Remark": "",
                }
            ]
        )

        merged = core._merge_coding_with_existing(existing_df, new_df)

        self.assertEqual(len(merged), 1)
        self.assertEqual(merged.iloc[0]["Question"], "s11_3_O4")

    def test_build_coding_df_skips_ambiguous_fallback_when_multiple_oth_have_text(self):
        df = pd.DataFrame(
            [
                {
                    "Sbjnum": "230066030",
                    "s14_1_O2": "93",
                    "s14_1_O4": "",
                    "s14_1_O6": "",
                    "s14_1_O8": "",
                    "s14_1_O10": "",
                    "s14_1_94_oth": "fgn",
                    "s14_1_95_oth": "sg",
                }
            ]
        )
        variable_labels = {
            "s14_1_O2": "Question 2",
            "s14_1_O4": "Question 4",
            "s14_1_O6": "Question 6",
            "s14_1_O8": "Question 8",
            "s14_1_O10": "Question 10",
        }
        other_vl = {
            93: "อื่นๆ ระบุ",
            94: "อื่นๆ ระบุ",
            95: "อื่นๆ ระบุ",
        }
        value_labels = {col: other_vl for col in variable_labels}

        coding_df, _ = core.build_coding_df(df, [], variable_labels, value_labels)

        self.assertTrue(coding_df.empty)

    def test_save_and_read_coding_sheet_roundtrip_preserves_rows_across_sheets(self):
        coding_df = pd.DataFrame(
            [
                {
                    "Question": "s6",
                    "Variable_Label": "Question 6",
                    core.SBJNUM_COL: "1",
                    "Other_Label": "อื่นๆ ระบุ",
                    "Other_Code": "97",
                    core.NEW_CODE_COL: "",
                    "Open_Text": "ตอบ 1",
                    core.NEW_OPEN_TEXT_COL: "",
                    "Open_Text_From": "s6_oth",
                    "Remark": "",
                },
                {
                    "Question": "s10_1_O1",
                    "Variable_Label": "Question 10_1",
                    core.SBJNUM_COL: "2",
                    "Other_Label": "อื่นๆ ระบุ",
                    "Other_Code": "98",
                    core.NEW_CODE_COL: "",
                    "Open_Text": "ตอบ 2",
                    core.NEW_OPEN_TEXT_COL: "",
                    "Open_Text_From": "s10_1_98_oth",
                    "Remark": "",
                },
            ]
        )

        tmp_path = self._workspace_tempdir()
        try:
            out_path = tmp_path / "codesheet.xlsx"
            core.save_coding_sheet(
                coding_df,
                out_path,
                source_columns=["Sbjnum", "s6", "s6_oth", "s10_1_O1", "s10_1_98_oth"],
            )
            read_back = core._read_existing_coding_sheet(out_path)
        finally:
            for file_path in tmp_path.glob("*"):
                try:
                    file_path.unlink(missing_ok=True)
                except PermissionError:
                    pass
            try:
                tmp_path.rmdir()
            except OSError:
                pass

        self.assertEqual(len(read_back), 2)
        self.assertEqual(set(read_back["Question"]), {"s6", "s10_1_O1"})

    def test_phase2_apply_new_code_updates_question_only(self):
        raw_df = pd.DataFrame(
            [
                {"Sbjnum": "230066030", "s6": "97", "s6_oth": "เดิม"},
                {"Sbjnum": "dummy", "s6": "keep-as-text", "s6_oth": ""},
            ]
        )
        coding_df = pd.DataFrame(
            [
                {
                    "Question": "s6",
                    "Variable_Label": "Question 6",
                    core.SBJNUM_COL: "230066030",
                    "Other_Label": "อื่นๆ ระบุ",
                    "Other_Code": "97",
                    core.NEW_CODE_COL: "500",
                    "Open_Text": "เดิม",
                    core.NEW_OPEN_TEXT_COL: "",
                    "Open_Text_From": "s6_oth",
                    "Remark": "",
                }
            ]
        )

        tmp_path = self._workspace_tempdir()
        try:
            raw_path = tmp_path / "raw.xlsx"
            coding_path = tmp_path / "coding.xlsx"
            out_path = tmp_path / "out.xlsx"
            raw_df.to_excel(raw_path, index=False)
            coding_df.to_excel(coding_path, index=False)
            core.phase2_apply(raw_path, coding_path, out_path)
            applied_df = pd.read_excel(out_path, dtype=str).fillna("")
        finally:
            for file_path in tmp_path.glob("*"):
                try:
                    file_path.unlink(missing_ok=True)
                except PermissionError:
                    pass
            try:
                tmp_path.rmdir()
            except OSError:
                pass

        self.assertEqual(applied_df.at[0, "s6"], "500")
        self.assertEqual(applied_df.at[0, "s6_oth"], "เดิม")

    def test_phase2_apply_new_code_cut_clears_question_only(self):
        raw_df = pd.DataFrame(
            [
                {"Sbjnum": "230066030", "s6": "97", "s6_oth": "เดิม"},
                {"Sbjnum": "dummy", "s6": "keep-as-text", "s6_oth": ""},
            ]
        )
        coding_df = pd.DataFrame(
            [
                {
                    "Question": "s6",
                    "Variable_Label": "Question 6",
                    core.SBJNUM_COL: "230066030",
                    "Other_Label": "อื่นๆ ระบุ",
                    "Other_Code": "97",
                    core.NEW_CODE_COL: "ตัด",
                    "Open_Text": "เดิม",
                    core.NEW_OPEN_TEXT_COL: "",
                    "Open_Text_From": "s6_oth",
                    "Remark": "",
                }
            ]
        )

        tmp_path = self._workspace_tempdir()
        try:
            raw_path = tmp_path / "raw.xlsx"
            coding_path = tmp_path / "coding.xlsx"
            out_path = tmp_path / "out.xlsx"
            raw_df.to_excel(raw_path, index=False)
            coding_df.to_excel(coding_path, index=False)
            core.phase2_apply(raw_path, coding_path, out_path)
            applied_df = pd.read_excel(out_path, dtype=str).fillna("")
        finally:
            for file_path in tmp_path.glob("*"):
                try:
                    file_path.unlink(missing_ok=True)
                except PermissionError:
                    pass
            try:
                tmp_path.rmdir()
            except OSError:
                pass

        self.assertEqual(applied_df.at[0, "s6"], "")
        self.assertEqual(applied_df.at[0, "s6_oth"], "เดิม")

    def test_phase2_apply_new_open_text_updates_verbatim_only(self):
        raw_df = pd.DataFrame(
            [
                {"Sbjnum": "230066030", "s6": "97", "s6_oth": "เดิม"},
                {"Sbjnum": "dummy", "s6": "keep-as-text", "s6_oth": ""},
            ]
        )
        coding_df = pd.DataFrame(
            [
                {
                    "Question": "s6",
                    "Variable_Label": "Question 6",
                    core.SBJNUM_COL: "230066030",
                    "Other_Label": "อื่นๆ ระบุ",
                    "Other_Code": "97",
                    core.NEW_CODE_COL: "",
                    "Open_Text": "เดิม",
                    core.NEW_OPEN_TEXT_COL: "แก้แล้ว",
                    "Open_Text_From": "s6_oth",
                    "Remark": "",
                }
            ]
        )

        tmp_path = self._workspace_tempdir()
        try:
            raw_path = tmp_path / "raw.xlsx"
            coding_path = tmp_path / "coding.xlsx"
            out_path = tmp_path / "out.xlsx"
            raw_df.to_excel(raw_path, index=False)
            coding_df.to_excel(coding_path, index=False)
            core.phase2_apply(raw_path, coding_path, out_path)
            applied_df = pd.read_excel(out_path, dtype=str).fillna("")
        finally:
            for file_path in tmp_path.glob("*"):
                try:
                    file_path.unlink(missing_ok=True)
                except PermissionError:
                    pass
            try:
                tmp_path.rmdir()
            except OSError:
                pass

        self.assertEqual(applied_df.at[0, "s6"], "97")
        self.assertEqual(applied_df.at[0, "s6_oth"], "แก้แล้ว")

    def test_phase2_apply_cut_in_new_open_text_clears_both_verbatim_and_matching_code(self):
        raw_df = pd.DataFrame(
            [
                {
                    "Sbjnum": "230066030",
                    "s14_1_O2": "93",
                    "s14_1_93_oth": "fhm hg",
                },
                {
                    "Sbjnum": "dummy",
                    "s14_1_O2": "keep-as-text",
                    "s14_1_93_oth": "",
                },
            ]
        )
        coding_df = pd.DataFrame(
            [
                {
                    "Question": "s14_1_O2",
                    "Variable_Label": "Question 2",
                    core.SBJNUM_COL: "230066030",
                    "Other_Label": "อื่นๆ ระบุ",
                    "Other_Code": "93",
                    core.NEW_CODE_COL: "",
                    "Open_Text": "fhm hg",
                    core.NEW_OPEN_TEXT_COL: "ตัด",
                    "Open_Text_From": "s14_1_93_oth",
                    "Remark": "",
                }
            ]
        )

        tmp_path = self._workspace_tempdir()
        try:
            raw_path = tmp_path / "raw.xlsx"
            coding_path = tmp_path / "coding.xlsx"
            out_path = tmp_path / "out.xlsx"
            raw_df.to_excel(raw_path, index=False)
            coding_df.to_excel(coding_path, index=False)

            result = core.phase2_apply(raw_path, coding_path, out_path)
            applied_df = pd.read_excel(out_path, dtype=str).fillna("")
        finally:
            for file_path in tmp_path.glob("*"):
                try:
                    file_path.unlink(missing_ok=True)
                except PermissionError:
                    pass
            try:
                tmp_path.rmdir()
            except OSError:
                pass

        self.assertEqual(result.n_applied, 1)
        self.assertEqual(applied_df.at[0, "s14_1_O2"], "")
        self.assertEqual(applied_df.at[0, "s14_1_93_oth"], "")

    def test_phase2_apply_reports_sbjnum_not_found(self):
        raw_df = pd.DataFrame([{"Sbjnum": "230066030", "s6": "97", "s6_oth": "เดิม"}])
        coding_df = pd.DataFrame(
            [
                {
                    "Question": "s6",
                    "Variable_Label": "Question 6",
                    core.SBJNUM_COL: "9999",
                    "Other_Label": "อื่นๆ ระบุ",
                    "Other_Code": "97",
                    core.NEW_CODE_COL: "500",
                    "Open_Text": "เดิม",
                    core.NEW_OPEN_TEXT_COL: "",
                    "Open_Text_From": "s6_oth",
                    "Remark": "",
                }
            ]
        )

        tmp_path = self._workspace_tempdir()
        try:
            raw_path = tmp_path / "raw.xlsx"
            coding_path = tmp_path / "coding.xlsx"
            out_path = tmp_path / "out.xlsx"
            raw_df.to_excel(raw_path, index=False)
            coding_df.to_excel(coding_path, index=False)
            result = core.phase2_apply(raw_path, coding_path, out_path)
        finally:
            for file_path in tmp_path.glob("*"):
                try:
                    file_path.unlink(missing_ok=True)
                except PermissionError:
                    pass
            try:
                tmp_path.rmdir()
            except OSError:
                pass

        self.assertEqual(result.n_applied, 0)
        self.assertEqual(result.n_not_found, 1)
        self.assertEqual(result.not_found_df.iloc[0]["Reason"], "sbjnum not found")

    def test_phase2_apply_reports_open_text_target_not_found(self):
        raw_df = pd.DataFrame(
            [
                {"Sbjnum": "230066030", "unknown_q": "97"},
                {"Sbjnum": "dummy", "unknown_q": "keep-as-text"},
            ]
        )
        coding_df = pd.DataFrame(
            [
                {
                    "Question": "unknown_q",
                    "Variable_Label": "Unknown Question",
                    core.SBJNUM_COL: "230066030",
                    "Other_Label": "อื่นๆ ระบุ",
                    "Other_Code": "97",
                    core.NEW_CODE_COL: "",
                    "Open_Text": "เดิม",
                    core.NEW_OPEN_TEXT_COL: "แก้แล้ว",
                    "Open_Text_From": "missing_oth_col",
                    "Remark": "",
                }
            ]
        )

        tmp_path = self._workspace_tempdir()
        try:
            raw_path = tmp_path / "raw.xlsx"
            coding_path = tmp_path / "coding.xlsx"
            out_path = tmp_path / "out.xlsx"
            raw_df.to_excel(raw_path, index=False)
            coding_df.to_excel(coding_path, index=False)
            result = core.phase2_apply(raw_path, coding_path, out_path)
        finally:
            for file_path in tmp_path.glob("*"):
                try:
                    file_path.unlink(missing_ok=True)
                except PermissionError:
                    pass
            try:
                tmp_path.rmdir()
            except OSError:
                pass

        self.assertEqual(result.n_applied, 0)
        self.assertEqual(result.n_not_found, 1)
        self.assertEqual(
            result.not_found_df.iloc[0]["Reason"],
            "open_text target column not found",
        )

    def test_phase1_export_integration_writes_codesheet_with_mocked_spss(self):
        raw_df = pd.DataFrame(
            [
                {"Sbjnum": "230066030", "s6": "97", "s6_oth": "เดิม"},
                {"Sbjnum": "dummy", "s6": "", "s6_oth": ""},
            ]
        )

        tmp_path = self._workspace_tempdir()
        try:
            raw_path = tmp_path / "raw.xlsx"
            out_path = tmp_path / "CodeSheet.xlsx"
            raw_df.to_excel(raw_path, index=False)

            with patch.object(
                core,
                "read_spss_labels",
                return_value=(
                    {"s6": "Question 6"},
                    {"s6": {97: "อื่นๆ ระบุ"}},
                ),
            ):
                result = core.phase1_export(raw_path, tmp_path / "dummy.sav", out_path)
        finally:
            for file_path in tmp_path.glob("*"):
                try:
                    file_path.unlink(missing_ok=True)
                except PermissionError:
                    pass
            try:
                tmp_path.rmdir()
            except OSError:
                pass

        self.assertTrue(out_path.exists() or result.output_path == out_path)
        self.assertEqual(result.n_rows, 1)
        self.assertEqual(result.coding_df.iloc[0]["Question"], "s6")
        self.assertEqual(result.coding_df.iloc[0]["Other_Code"], "97")

    def test_phase1_export_supports_order_fallback_mapping(self):
        raw_df = pd.DataFrame(
            [
                {"Sbjnum": "230066030", "เนสกาแฟ_s10": "97", "เนสกาแฟ_s10_oth": "เดิม"},
                {"Sbjnum": "dummy", "เนสกาแฟ_s10": "", "เนสกาแฟ_s10_oth": ""},
            ]
        )

        tmp_path = self._workspace_tempdir()
        try:
            raw_path = tmp_path / "raw.xlsx"
            out_path = tmp_path / "CodeSheet.xlsx"
            raw_df.to_excel(raw_path, index=False)

            with patch.object(
                core,
                "read_spss_labels",
                return_value=(
                    {"l_1_s10": "Question s10"},
                    {"l_1_s10": {97: "อื่นๆ ระบุ"}},
                ),
            ):
                report = core.inspect_phase1_column_mapping(raw_path, tmp_path / "dummy.sav")
                self.assertEqual(report.unresolved_excel_cols, [])
                self.assertEqual(report.direct_matches, [("เนสกาแฟ_s10", "l_1_s10")])

                result = core.phase1_export(
                    raw_path,
                    tmp_path / "dummy.sav",
                    out_path,
                    allow_order_fallback=True,
                )
        finally:
            for file_path in tmp_path.glob("*"):
                try:
                    file_path.unlink(missing_ok=True)
                except PermissionError:
                    pass
            try:
                tmp_path.rmdir()
            except OSError:
                pass

        self.assertEqual(result.n_rows, 1)
        self.assertEqual(result.coding_df.iloc[0]["Question"], "เนสกาแฟ_s10")
        self.assertEqual(result.coding_df.iloc[0]["Variable_Label"], "Question s10")
        self.assertEqual(result.coding_df.iloc[0]["Other_Code"], "97")


if __name__ == "__main__":
    unittest.main()
