# score_filter_core.py
# -*- coding: utf-8 -*-
from pathlib import Path
import pandas as pd
import traceback

# 显式导入，帮助 PyInstaller 收集 openpyxl 子模块
try:
    import openpyxl  # noqa: F401
except Exception:
    # 不终止，仅用于打包收集依赖
    pass

DEFAULT_PUBCLASS_QUALIFIED_NUM = 10

def find_col_exact(columns, wanted):
    for c in columns:
        if str(c) == wanted:
            return c
    return None

def process_file(infile_path, pubclass_qualified_num=DEFAULT_PUBCLASS_QUALIFIED_NUM,
                 divide_output=False, log_fn=None):
    """
    处理主逻辑 — 返回 (success_bool, message_or_exception)
    log_fn: 可选回调，用于写日志（接收 str）
    """
    try:
        def log(s=""):
            if log_fn:
                log_fn(s)
            else:
                print(s)

        infile = str(infile_path)
        log(f"读取文件: {infile}")
        df = pd.read_excel(infile, dtype={0: str}, engine="openpyxl")
        cols = list(df.columns)

        # 关键列
        student_id_col = cols[0]
        course_type_col = find_col_exact(cols, "一层节点")
        course_name_col = find_col_exact(cols, "课程名称")
        credit_col      = find_col_exact(cols, "获得学分")
        score_col       = find_col_exact(cols, "成绩")
        log(f"列识别: 学号={student_id_col}, 成绩={score_col}, 学分={credit_col}")

        # 过滤“一层节点=其它”
        if course_type_col is not None:
            before = len(df)
            df = df.loc[df[course_type_col].astype(str) != "其它"].copy()
            log(f"已过滤 '其它' 行：{before} -> {len(df)}")

        # 数值化
        for c in [score_col, credit_col]:
            if c in df.columns:
                df[c] = pd.to_numeric(df[c], errors="coerce")

        # 班级人数（去重学号）
        class_size = df[student_id_col].nunique(dropna=True)
        log(f"班级总人数: {class_size}")

        # 公选/非公选
        if course_type_col is not None:
            is_public = df[course_type_col].astype(str).str.contains("公共选修课", na=False)
        else:
            is_public = pd.Series(False, index=df.index)

        # 规则一：公选 & 学分<阈值 & 成绩<60/空
        if credit_col is not None and score_col is not None:
            df_public_lt = df.loc[
                is_public &
                (df[credit_col] < pubclass_qualified_num) &
                ((df[score_col] < 60) | (df[score_col].isna()))
            ].copy()
        else:
            df_public_lt = pd.DataFrame(columns=df.columns)
        public_count = df_public_lt[student_id_col].nunique()
        log(f"规则一：学生数 {public_count}，记录 {len(df_public_lt)}")

        # 规则二：非公选 & 成绩为空/0（需关注；若全班该课空/0，则视作未开设而不导出）
        not_public = ~is_public
        if score_col is not None:
            zero_like = df[score_col].isna() | (df[score_col] == 0)
            unpub_zero_rows = df.loc[not_public & zero_like].copy()
        else:
            unpub_zero_rows = pd.DataFrame(columns=df.columns)
        export_zero_rows = pd.DataFrame(columns=df.columns)
        not_offered_names = []

        if not unpub_zero_rows.empty and (course_name_col is not None):
            zero_counts = unpub_zero_rows.groupby(course_name_col)[student_id_col].nunique()
            for course, cnt in zero_counts.items():
                if cnt == class_size:
                    not_offered_names.append(course)  # 全员空/0 -> 未开设
                else:
                    export_zero_rows = pd.concat(
                        [export_zero_rows, unpub_zero_rows[unpub_zero_rows[course_name_col] == course]],
                        axis=0
                    )
        else:
            export_zero_rows = unpub_zero_rows.copy()

        log(f"规则二：需关注记录 {len(export_zero_rows)}，未开设课程 {len(not_offered_names)}")

        # 规则三：非公选 & 非空/0 & 成绩<60
        if score_col is not None:
            remaining = df.loc[not_public & ~(df[score_col].isna() | (df[score_col] == 0))].copy()
            fail_rows = remaining.loc[(remaining[score_col] < 60)].copy()
        else:
            fail_rows = pd.DataFrame(columns=df.columns)
        log(f"规则三：不及格人数 {fail_rows[student_id_col].nunique()}，记录 {len(fail_rows)}")

        # 标签
        if not df_public_lt.empty:
            df_public_lt['来源规则'] = f'规则一: 公选 学分<{pubclass_qualified_num} 且 成绩<60/空'
        if not export_zero_rows.empty:
            export_zero_rows['来源规则'] = '规则二: 非公选 正常开设 成绩0分/空白'
        if not fail_rows.empty:
            fail_rows['来源规则'] = '规则三: 非公选 正常开设 0<成绩<60'

        # 合并 + 去重
        df_final = pd.concat([export_zero_rows, fail_rows, df_public_lt], ignore_index=True)

        def _maybe_get_term_col(df_):
            for name in ['学年学期', '学期', '建议修读学年']:
                if name in df_.columns:
                    return name
            return None

        subset_cols = [c for c in [student_id_col, course_name_col, _maybe_get_term_col(df)] if c in df_final.columns and c]
        if subset_cols:
            before = len(df_final)
            df_final = df_final.drop_duplicates(subset=subset_cols, keep='first')
            log(f"去重 {subset_cols}: {before} -> {len(df_final)}")

        # 输出
        out_dir = Path(infile).parent
        out_xlsx_final = out_dir / f"学业预警表_{Path(infile).stem}.xlsx"
        df_final.to_excel(out_xlsx_final, index=False)
        log(f"总表已导出: {out_xlsx_final}")

        if divide_output:
            if not df_public_lt.empty:
                df_public_lt.to_csv(out_dir / f"{Path(infile).stem}_公选_学分小于{pubclass_qualified_num}.csv",
                                    index=False, encoding="utf-8-sig")
            if not export_zero_rows.empty:
                export_zero_rows.to_csv(out_dir / f"{Path(infile).stem}_其他_0分需关注.csv",
                                        index=False, encoding="utf-8-sig")
            if not fail_rows.empty:
                fail_rows.to_csv(out_dir / f"{Path(infile).stem}_其他_不及格小于60.csv",
                                 index=False, encoding="utf-8-sig")
            log("分规则 CSV 已输出（divide_output=True）")

        summary = (
            f"班级总人数: {class_size}\n"
            f"规则一: 学分<{pubclass_qualified_num} 且 成绩<60/空 - 学生数 {public_count}, 记录 {len(df_public_lt)}\n"
            f"规则二: 需关注记录 {len(export_zero_rows)}, 未开设课程数 {len(not_offered_names)}\n"
            f"规则三: 不及格人数 {fail_rows[student_id_col].nunique()}, 记录 {len(fail_rows)}\n"
            f"输出文件: {out_xlsx_final}"
        )
        log(summary)
        return True, summary

    except Exception:
        tb = traceback.format_exc()
        if log_fn:
            log_fn(tb)
        return False, tb
