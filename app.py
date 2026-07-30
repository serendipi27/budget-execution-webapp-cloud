"""
예산 집행 취합·검증·공문 초안 자동화 — Streamlit 웹앱.

vibecoding.py에 옮겨둔 vibecoding.ipynb Step 1~9의 실제 코드를 그대로 호출한다 —
새로 만든 로직이 아니라 이미 검증된 코드를 재사용하는 것이 목적이다.
"""

import contextlib
import io
import os
import re

import streamlit as st
import pandas as pd

import vibecoding

st.set_page_config(page_title="예산 집행 자동화", page_icon="📊", layout="wide")

MASTER_PATH_DEFAULT = "webapp_data/2026년도_예산집행취합.xlsx"


def derived_dir(name):
    """마스터 파일과 같은 폴더 밑에 있는 하위 폴더 경로(그래프/이메일/공문자료 등)를 계산한다."""
    return os.path.join(os.path.dirname(st.session_state.master_path) or ".", name)


_MD_IMAGE_RE = re.compile(r"!\[([^\]]*)\]\(([^)]+)\)")


def render_markdown_with_local_images(markdown_text, base_dir):
    """st.markdown()은 로컬 파일의 상대경로 이미지(예: ![...](그래프/집행률_비교_2026-07.png))를
    표시하지 못한다 — Streamlit이 그런 경로를 정적 파일로 서빙해주지 않기 때문이다.
    대신 텍스트를 이미지 태그 기준으로 잘라, 텍스트는 st.markdown()으로, 이미지는
    matplotlib이 디스크에 저장해둔 PNG 파일을 st.image()로 직접 읽어 보여준다 —
    원래 순서 그대로 번갈아 렌더링된다(vibecoding.py가 만드는 md 파일 자체는 건드리지 않음).

    LLM이 쓰는 이미지 경로의 디렉터리 부분은 신뢰하지 않는다 — Step 8-1 원자료의
    "파일 위치" 문구를 그대로 베껴서 이미 base_dir이 포함된 전체 경로를 쓰는 경우가
    있어(예: `webapp_data/그래프/...`), base_dir과 합치면 두 번 겹쳐서 실제로 존재하는
    파일인데도 못 찾는 문제가 있었다. 그래서 파일명만 뽑아 실제 그래프 폴더에서
    찾는 방식을 우선한다."""
    graph_dir = os.path.join(base_dir, "그래프")
    last_end = 0
    for m in _MD_IMAGE_RE.finditer(markdown_text):
        if m.start() > last_end:
            st.markdown(markdown_text[last_end:m.start()])
        alt, path = m.group(1), m.group(2)
        candidates = [
            os.path.join(graph_dir, os.path.basename(path)),  # 파일명만으로 실제 그래프 폴더에서 찾기 (우선)
            path if os.path.isabs(path) else os.path.join(base_dir, path),  # LLM이 쓴 경로 그대로 (그다음)
        ]
        full_path = next((p for p in candidates if os.path.exists(p)), None)
        if full_path:
            st.image(full_path, caption=alt or None)
        else:
            st.caption(f"⚠️ 이미지를 찾을 수 없습니다: {os.path.basename(path)}")
        last_end = m.end()
    if last_end < len(markdown_text):
        st.markdown(markdown_text[last_end:])


# ===========================================================================
# 파이프라인 단계 정의 (vibecoding.ipynb Step 1~11 매핑, UI용으로 8단계로 묶음)
# ===========================================================================

STAGES = [
    {"key": "collect", "icon": "📥", "label": "데이터 취합", "steps": "Step 1"},
    {"key": "validate", "icon": "🔍", "label": "데이터 검증", "steps": "Step 2"},
    {"key": "issue", "icon": "✉️", "label": "이슈 확인 및 정정", "steps": "Step 3-4"},
    {"key": "revalidate", "icon": "🔁", "label": "재취합·재검증", "steps": "Step 5"},
    {"key": "metrics", "icon": "📊", "label": "지표 및 그래프", "steps": "Step 6-7"},
    {"key": "prepare", "icon": "📝", "label": "보고 자료 준비", "steps": "Step 8"},
    {"key": "draft", "icon": "🤖", "label": "보고서 초안", "steps": "Step 9"},
    {"key": "review", "icon": "✅", "label": "검토 및 저장", "steps": "Step 10-11"},
]

STATUS_ICON = {"done": "✅", "in_progress": "🔄", "not_started": "⬜"}


def init_state():
    if "master_path" not in st.session_state:
        st.session_state.master_path = MASTER_PATH_DEFAULT

    if "stage_status" not in st.session_state:
        st.session_state.stage_status = {s["key"]: "not_started" for s in STAGES}

    if "current_stage" not in st.session_state:
        st.session_state.current_stage = "collect"

    if "merge_log_text" not in st.session_state:
        st.session_state.merge_log_text = ""

    if "validation_log_text" not in st.session_state:
        st.session_state.validation_log_text = ""

    if "issues" not in st.session_state:
        st.session_state.issues = {}  # id -> issue dict

    if "revalidate_log_text" not in st.session_state:
        st.session_state.revalidate_log_text = ""

    if "acknowledged_issues" not in st.session_state:
        st.session_state.acknowledged_issues = set()  # 해결이 어렵다고 사람이 인지·확인한 issue id

    if "review_saved_path" not in st.session_state:
        st.session_state.review_saved_path = ""

    if "draft_previews" not in st.session_state:
        st.session_state.draft_previews = {}  # 연도-월 -> LLM이 쓴 초안 텍스트 (아직 디스크에 저장 전)


def set_stage(key):
    st.session_state.current_stage = key


def advance_to(stage_key, done_key):
    """이전 단계를 '완료'로 표시하면서 다음 화면으로 이동한다(사이드바 체크 표시용)."""
    st.session_state.stage_status[done_key] = "done"
    st.session_state.current_stage = stage_key


def run_captured(fn, *args, **kwargs):
    """fn 실행 중 print() 출력을 문자열로 캡처해서 (결과, 로그텍스트)로 반환한다."""
    buf = io.StringIO()
    with contextlib.redirect_stdout(buf):
        result = fn(*args, **kwargs)
    return result, buf.getvalue()


# ===========================================================================
# 사이드바
# ===========================================================================

def render_sidebar():
    st.sidebar.title("📊 예산 집행 보고서 작성 자동화")
    st.sidebar.caption("월간 파이프라인 진행 현황")

    done_count = sum(1 for v in st.session_state.stage_status.values() if v == "done")
    st.sidebar.progress(done_count / len(STAGES))
    st.sidebar.caption(f"{done_count} / {len(STAGES)} 단계 완료")

    st.sidebar.divider()

    for stage in STAGES:
        status = st.session_state.stage_status[stage["key"]]
        label = f"{STATUS_ICON[status]} {stage['icon']} {stage['label']}"
        is_current = st.session_state.current_stage == stage["key"]
        st.sidebar.button(
            label,
            key=f"nav_{stage['key']}",
            use_container_width=True,
            type="primary" if is_current else "secondary",
            on_click=set_stage,
            args=(stage["key"],),
        )

    st.sidebar.divider()
    months = vibecoding.list_data_months(st.session_state.master_path)
    st.sidebar.caption(f"취합된 월: **{', '.join(months) if months else '없음'}**")
    with st.sidebar.expander("⚙️ 설정"):
        st.session_state.master_path = st.text_input("마스터 파일 경로", value=st.session_state.master_path)


# ===========================================================================
# Step 1 — 데이터 취합 (vibecoding.run_merge_pipeline 그대로 호출)
# ===========================================================================

def page_collect():
    st.header("📥 데이터 취합")
    st.caption("Step 1 — 부서별 제출 파일을 하나의 마스터 파일로 통합합니다.")
    st.info(
        "루트 폴더를 업로드하면, 그 안의 `연도월_부서제출파일` 형식 하위 폴더를 자동으로 "
        "구분해서 폴더별로(=월별로) 각각 취합합니다. 여러 폴더가 함께 있어도 한 번에 처리됩니다.",
        icon="ℹ️",
    )

    uploaded = st.file_uploader(
        "부서 제출 파일이 들어있는 루트 폴더 업로드",
        type=["xlsx", "csv"],
        accept_multiple_files="directory",
        help="폴더를 선택하면 하위 폴더(예: 202607_부서제출파일)를 포함해 전체가 업로드됩니다. "
             "xlsx·csv 외 파일(.DS_Store 등)은 자동으로 제외됩니다.",
    )

    if uploaded:
        supported, unsupported = vibecoding.filter_supported_files(uploaded)
        groups, ungrouped = vibecoding.group_uploads_by_folder(supported)
        if groups:
            found = ", ".join(
                f"{name}({len(files)}개)→{vibecoding.parse_sheet_name(name) or '인식불가'}"
                for name, files in sorted(groups.items())
            )
            st.caption(f"발견된 하위 폴더: {found}")
        if unsupported:
            st.caption(f"xlsx/csv가 아니어서 제외됨: {len(unsupported)}개 ({', '.join(unsupported[:5])}{' 등' if len(unsupported) > 5 else ''})")
        if ungrouped:
            st.warning(
                f"하위 폴더 없이 올라온 파일 {len(ungrouped)}개는 어느 달인지 알 수 없어 건너뜁니다: "
                + ", ".join(ungrouped),
                icon="⚠️",
            )

    if st.button("취합 실행", type="primary", disabled=not uploaded):
        with st.spinner("취합 중..."):
            months_before = set(vibecoding.list_data_months(st.session_state.master_path))
            _, log_text = run_captured(
                vibecoding.run_merge_pipeline_from_uploads, uploaded, st.session_state.master_path
            )
            months_after = set(vibecoding.list_data_months(st.session_state.master_path))
        st.session_state.merge_log_text = log_text
        if months_after > months_before:
            st.session_state.stage_status["collect"] = "done"
        st.rerun()

    if st.session_state.merge_log_text:
        st.subheader("실행 로그")
        st.code(st.session_state.merge_log_text, language=None)

    months = vibecoding.list_data_months(st.session_state.master_path)
    if months:
        st.subheader("현재 마스터 파일에 취합된 월")
        st.dataframe(pd.DataFrame({"연도-월": months}), use_container_width=True, hide_index=True)

        selected_month = st.selectbox("병합된 데이터 미리보기 — 연월 선택", months, index=len(months) - 1)
        rows = vibecoding.load_merged_rows(st.session_state.master_path, selected_month)
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

        st.button("다음: 데이터 검증 →", type="primary", on_click=set_stage, args=("validate",))


# ===========================================================================
# Step 2 — 데이터 검증 (vibecoding.run_validation_pipeline 그대로 호출)
# ===========================================================================

def sync_issues_from_validation():
    """이미 취합된 모든 달을 다시 validate()로 조회해 이슈 카드 목록과 동기화한다.
    새로 생긴 이슈는 추가하고, 더 이상 검증에 걸리지 않는(=실제로 고쳐진) 이슈는 카드
    자체를 제거한다 — '수정 완료'는 사람이 버튼으로 선언하는 게 아니라, 재취합·재검증
    결과가 실제로 그렇다는 것으로만 판단한다."""
    current_ids = set()
    for result in vibecoding.get_all_validation_results(st.session_state.master_path):
        month = result["month"]
        for kind, rows in (("집행액 누락", result["missing_rows"]), ("이상치(집행액 > 배정액)", result["outlier_rows"])):
            for r in rows:
                issue_id = f"{month}|{r['사업장']}|{r['부서명']}|{r['예산과목']}|{kind}"
                current_ids.add(issue_id)
                if issue_id not in st.session_state.issues:
                    st.session_state.issues[issue_id] = {
                        "id": issue_id,
                        "월": month,
                        "부서": f"{r['사업장']} {r['부서명']}",
                        "예산과목": r["예산과목"],
                        "유형": kind,
                        "담당자": r["작성자"] or "(미상)",
                        "배정액": r["배정액"],
                        "집행액": r["집행액"],
                        "상태": "이슈없음",
                        "email_draft": None,
                        "reply_text": "",
                    }

    for stale_id in list(st.session_state.issues.keys()):
        if stale_id not in current_ids:
            del st.session_state.issues[stale_id]


def page_validate():
    st.header("🔍 데이터 검증")
    st.caption("Step 2 — 미제출 부서 / 집행액 누락 / 이상치를 Rule 기반으로 탐지합니다.")

    months = vibecoding.list_data_months(st.session_state.master_path)
    if not months:
        st.warning("먼저 데이터 취합을 실행해 주세요.", icon="⚠️")
        return

    unvalidated = vibecoding.find_unvalidated_sheets(st.session_state.master_path)
    st.caption(f"검증 대상: {', '.join(unvalidated) if unvalidated else '없음(모두 검증됨)'}")

    if st.button("검증 실행", type="primary", disabled=not unvalidated):
        with st.spinner("검증 중..."):
            _, log_text = run_captured(vibecoding.run_validation_pipeline, st.session_state.master_path)
        st.session_state.validation_log_text = log_text
        sync_issues_from_validation()
        st.session_state.stage_status["validate"] = "done"
        st.rerun()

    if st.session_state.validation_log_text:
        st.subheader("실행 로그")
        st.code(st.session_state.validation_log_text, language=None)

    results = vibecoding.get_all_validation_results(st.session_state.master_path)
    if results:
        st.subheader("검증 결과 (마스터 파일 기준 전체)")
        total_missing_dept = sum(len(r["missing_depts"]) for r in results)
        total_missing_row = sum(len(r["missing_rows"]) for r in results)
        total_outlier = sum(len(r["outlier_rows"]) for r in results)
        c1, c2, c3 = st.columns(3)
        c1.metric("미제출 부서", f"{total_missing_dept}건")
        c2.metric("집행액 누락", f"{total_missing_row}건")
        c3.metric("이상치", f"{total_outlier}건")

        for result in results:
            with st.expander(f"{result['month']} 상세 결과", expanded=False):
                rows = []
                for 사업장, 부서명 in result["missing_depts"]:
                    rows.append({"구분": "미제출부서", "내용": f"{사업장} {부서명}"})
                for r in result["missing_rows"]:
                    rows.append({"구분": "누락값", "내용": f"{r['사업장']} {r['부서명']} / {r['예산과목']} 집행액 미기재"})
                for r in result["outlier_rows"]:
                    rows.append({
                        "구분": "이상치",
                        "내용": f"{r['사업장']} {r['부서명']} / {r['예산과목']} 배정액 {r['배정액']:,}원 < 집행액 {r['집행액']:,}원",
                    })
                if rows:
                    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
                else:
                    st.caption("이상 없음")

        st.button("다음: 이슈 확인 및 정정 →", type="primary", on_click=set_stage, args=("issue",))


# ===========================================================================
# Step 3-4 — 이슈 확인 및 정정 (이슈 목록은 실제 검증 결과 기반, 이메일/정정은 아직 목업)
# ===========================================================================

def render_issue_card(issue):
    status = issue["상태"]
    badge = {
        "이슈없음": "🟢",
        "이메일발송됨(회신대기)": "🟡",
        "회신접수": "🟠",
    }.get(status, "⬜")

    with st.container(border=True):
        top = st.columns([2, 2, 2, 2, 2])
        top[0].markdown(f"**{issue['월']}**")
        top[1].markdown(f"{issue['부서']} / {issue['예산과목']}")
        top[2].markdown(issue["유형"])
        top[3].markdown(f"담당자: {issue['담당자']}")
        top[4].markdown(f"{badge} {status}")

        tab1, tab2 = st.tabs(["📧 확인 이메일", "💬 담당자 회신"])

        with tab1:
            if issue["email_draft"] is None:
                if st.button("이메일 초안 생성 (Upstage LLM)", key=f"draft_{issue['id']}"):
                    with st.spinner("LLM으로 이메일 초안 작성 중..."):
                        사업장, 부서명 = issue["부서"].split(" ", 1)
                        # vibecoding.py의 save_validation_log()가 검증로그에 기록하는 것과
                        # 동일한 "[구분] 내용" 형식으로 맞춰서, 노트북의 draft_email()에
                        # 그대로 넘긴다 — group_issues_by_dept()가 실제로 만드는 것과 같은 입력.
                        if issue["유형"] == "집행액 누락":
                            issue_desc = f"[누락값] {issue['예산과목']} 집행액 미기재"
                        else:
                            issue_desc = (
                                f"[이상치] {issue['예산과목']} "
                                f"배정액 {issue['배정액']:,}원 < 집행액 {issue['집행액']:,}원"
                            )
                        issue["email_draft"] = vibecoding.draft_email(
                            사업장, 부서명, issue["담당자"], issue["월"], [issue_desc]
                        )
                    issue["상태"] = "이메일발송됨(회신대기)"
                    st.rerun()
            else:
                st.text_area("초안 내용", issue["email_draft"], height=140, key=f"draft_view_{issue['id']}")
                st.download_button(
                    "다운로드", issue["email_draft"],
                    file_name=f"이메일_{issue['부서']}_{issue['월']}.txt", key=f"dl_{issue['id']}",
                )
                st.caption("실제 발송은 메일 클라이언트로 직접 진행해 주세요.")

        with tab2:
            reply = st.text_area(
                "담당자 회신 내용을 실제 이메일에서 확인 후 붙여넣어 주세요",
                value=issue.get("reply_text", ""), height=140, key=f"reply_{issue['id']}",
            )
            if st.button("회신 확인함", key=f"confirm_{issue['id']}"):
                issue["reply_text"] = reply
                issue["상태"] = "회신접수"
                # Step 8-1(find_reply_files/build_final_markdown)이 실제 이 폴더의
                # '이메일_응답_사업장_담당자_연도-월.txt' 파일을 읽어 보고서 원자료에
                # 포함시키므로, 세션에만 남기지 않고 노트북과 같은 파일명 규칙으로 저장한다.
                사업장, 부서명 = issue["부서"].split(" ", 1)
                email_dir = derived_dir("이메일")
                os.makedirs(email_dir, exist_ok=True)
                reply_path = os.path.join(email_dir, f"이메일_응답_{사업장}_{issue['담당자']}_{issue['월']}.txt")
                with open(reply_path, "w", encoding="utf-8") as f:
                    f.write(reply)
                st.rerun()

        사업장, 부서명 = issue["부서"].split(" ", 1)
        yyyymm = issue["월"].replace("-", "")
        filename = f"예산집행_{사업장}_{부서명}_{yyyymm}.xlsx"
        st.markdown(
            f"**정정 절차 (사람이 직접 수행 — Step 4-2)**\n\n"
            f"- 파일 복사(백업): `{filename}` → `original_{filename}`\n"
            f"- 파일 내용 수정: `{filename}` — 위 담당자 회신 내용을 참고해서 실제 값을 정정\n"
            f"- 수정이 끝나면 **'🔁 재취합·재검증' 화면에서 같은 루트 폴더를 다시 업로드하고 실행**하세요. "
            f"`original_` 백업 파일이 자동으로 감지되어 보관 폴더로 옮겨지고, "
            f"남은 정정본으로 재취합·재검증까지 이어집니다."
        )


def page_issue():
    st.header("✉️ 이슈 확인 및 정정")
    st.caption("Step 3-4 — 검증 이슈를 담당자에게 확인하고, 회신을 바탕으로 데이터를 정정합니다.")

    # 이 페이지를 열 때마다 마스터 파일의 실제 검증 결과와 다시 동기화한다.
    # (page_validate의 '검증 실행' 버튼은 새로 검증할 달이 없으면 비활성화되므로,
    # 그 버튼 클릭에만 의존하면 세션 상태가 초기화된 경우 — 예: 서버 재시작 —
    # 이미 파일에 저장된 이슈를 놓칠 수 있다.)
    sync_issues_from_validation()

    if not st.session_state.issues:
        st.success("현재 확인이 필요한 이슈가 없습니다.", icon="✅")
        st.button("다음: 재취합·재검증 →", type="primary", on_click=advance_to, args=("revalidate", "issue"))
        return

    st.info(
        "이 앱은 실제 이메일 발신/수신을 자동으로 처리하지 않습니다. "
        "초안 생성까지만 자동화하고, 발송과 회신 확인은 사람이 직접 수행한 뒤 그 결과를 아래에 반영해 주세요. "
        "실제 정정은 소스 파일을 고친 뒤 '🔁 재취합·재검증' 화면에서 반영합니다.",
        icon="ℹ️",
    )

    for issue in st.session_state.issues.values():
        render_issue_card(issue)

    st.button("다음: 재취합·재검증 →", type="primary", on_click=advance_to, args=("revalidate", "issue"))


# ===========================================================================
# Step 5 — 재취합·재검증 (vibecoding.run_reprocess_pipeline_from_uploads 그대로 호출)
# ===========================================================================

def page_revalidate():
    st.header("🔁 재취합·재검증")
    st.caption("Step 5 — 정정된 데이터를 다시 취합하고 검증합니다 (`original_` 백업 파일 자동 감지).")
    st.info(
        "실제 소스 파일을 정정(원본은 `original_` 접두어로 백업)한 뒤, 같은 루트 폴더를 다시 "
        "업로드하고 아래 '재취합·재검증 실행'을 눌러주세요. `original_` 백업 파일이 있는 달만 "
        "자동으로 골라 재처리합니다.",
        icon="ℹ️",
    )

    uploaded = st.file_uploader(
        "부서 제출 파일이 들어있는 루트 폴더 다시 업로드",
        type=["xlsx", "csv"],
        accept_multiple_files="directory",
        key="revalidate_uploader",
        help="정정한 파일과 `original_` 백업 파일이 그대로 들어있는 폴더를 업로드하세요.",
    )

    if uploaded:
        supported, unsupported = vibecoding.filter_supported_files(uploaded)
        groups, ungrouped = vibecoding.group_uploads_by_folder(supported)
        if groups:
            found = ", ".join(
                f"{name}({len(files)}개)→{vibecoding.parse_sheet_name(name) or '인식불가'}"
                for name, files in sorted(groups.items())
            )
            st.caption(f"발견된 하위 폴더: {found}")
        if unsupported:
            st.caption(f"xlsx/csv가 아니어서 제외됨: {len(unsupported)}개")
        if ungrouped:
            st.warning(
                f"하위 폴더 없이 올라온 파일 {len(ungrouped)}개는 어느 달인지 알 수 없어 건너뜁니다: "
                + ", ".join(ungrouped),
                icon="⚠️",
            )

    if st.button("🔁 재취합·재검증 실행", type="primary", disabled=not uploaded):
        with st.spinner("original 백업 파일 감지 및 재처리 중..."):
            reprocessed, log_text = run_captured(
                vibecoding.run_reprocess_pipeline_from_uploads, uploaded, st.session_state.master_path
            )
        st.session_state.revalidate_log_text = log_text
        sync_issues_from_validation()
        if reprocessed:
            st.session_state.stage_status["validate"] = "done"
        st.rerun()

    if st.session_state.revalidate_log_text:
        st.subheader("실행 로그")
        st.code(st.session_state.revalidate_log_text, language=None)

    months = vibecoding.list_data_months(st.session_state.master_path)
    if months:
        st.subheader("현재 마스터 파일에 취합된 월")
        st.dataframe(pd.DataFrame({"연도-월": months}), use_container_width=True, hide_index=True)

        selected_month = st.selectbox(
            "병합된 데이터 미리보기 — 연월 선택", months, index=len(months) - 1, key="revalidate_month_select"
        )
        rows = vibecoding.load_merged_rows(st.session_state.master_path, selected_month)
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

    results = vibecoding.get_all_validation_results(st.session_state.master_path)
    if results:
        st.subheader("검증 결과 (마스터 파일 기준 전체)")
        total_missing_dept = sum(len(r["missing_depts"]) for r in results)
        total_missing_row = sum(len(r["missing_rows"]) for r in results)
        total_outlier = sum(len(r["outlier_rows"]) for r in results)
        c1, c2, c3 = st.columns(3)
        c1.metric("미제출 부서", f"{total_missing_dept}건")
        c2.metric("집행액 누락", f"{total_missing_row}건")
        c3.metric("이상치", f"{total_outlier}건")

        for result in results:
            with st.expander(f"{result['month']} 상세 결과", expanded=False):
                rows = []
                for 사업장, 부서명 in result["missing_depts"]:
                    rows.append({"구분": "미제출부서", "내용": f"{사업장} {부서명}"})
                for r in result["missing_rows"]:
                    rows.append({"구분": "누락값", "내용": f"{r['사업장']} {r['부서명']} / {r['예산과목']} 집행액 미기재"})
                for r in result["outlier_rows"]:
                    rows.append({
                        "구분": "이상치",
                        "내용": f"{r['사업장']} {r['부서명']} / {r['예산과목']} 배정액 {r['배정액']:,}원 < 집행액 {r['집행액']:,}원",
                    })
                if rows:
                    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
                else:
                    st.caption("이상 없음")

    # 이 화면에 들어올 때마다 마스터 파일의 실제 상태와 이슈 목록을 다시 동기화한다 —
    # 재취합·재검증을 실행하지 않고 다른 경로로 들어왔을 때도 최신 상태를 보여주기 위함.
    sync_issues_from_validation()

    st.subheader("남은 이슈 처리 확인")
    if not st.session_state.issues:
        st.success("남은 이슈가 없습니다 — 정정이 모두 정상적으로 반영되었습니다.", icon="✅")
        can_proceed = True
    else:
        st.caption(
            "정정이 아직 반영되지 않았거나, 이번 달은 바로 해결하기 어려운 이슈는(예: 다음 달 소급 반영 "
            "예정) 아래에서 '인지함'을 체크하면 다음 단계로 진행할 수 있습니다."
        )
        for issue in st.session_state.issues.values():
            cols = st.columns([5, 2])
            cols[0].markdown(f"**{issue['월']}** · {issue['부서']} / {issue['예산과목']} · {issue['유형']}")
            checked = cols[1].checkbox(
                "인지함",
                key=f"ack_{issue['id']}",
                value=issue["id"] in st.session_state.acknowledged_issues,
            )
            if checked:
                st.session_state.acknowledged_issues.add(issue["id"])
            else:
                st.session_state.acknowledged_issues.discard(issue["id"])
        can_proceed = all(i["id"] in st.session_state.acknowledged_issues for i in st.session_state.issues.values())

    if st.button("다음: 지표 및 그래프 →", type="primary", disabled=not can_proceed):
        st.session_state.stage_status["revalidate"] = "done"
        set_stage("metrics")
        st.rerun()
    if not can_proceed:
        st.caption("모든 이슈가 해결되거나 '인지함' 체크가 되어야 다음 단계로 진행할 수 있습니다.")


def page_metrics():
    st.header("📊 지표 및 그래프")
    st.caption("Step 6-7 — 집행률·전월 대비 증감을 계산하고 그래프를 생성합니다.")

    months = vibecoding.list_data_months(st.session_state.master_path)
    if not months:
        st.warning("먼저 데이터 취합을 실행해 주세요.", icon="⚠️")
        return

    graph_dir = derived_dir("그래프")

    if st.button("지표 계산 및 그래프 생성", type="primary"):
        with st.spinner("지표 계산 및 그래프 생성 중..."):
            _, m_log = run_captured(vibecoding.run_metrics_pipeline, st.session_state.master_path)
            _, g_log = run_captured(vibecoding.run_graph_pipeline, st.session_state.master_path, graph_dir)
        st.session_state.metrics_log_text = m_log + "\n" + g_log
        st.session_state.stage_status["metrics"] = "done"
        st.rerun()

    if st.session_state.get("metrics_log_text"):
        st.subheader("실행 로그")
        st.code(st.session_state.metrics_log_text, language=None)

    metric_months = vibecoding.list_metric_months(st.session_state.master_path)
    if metric_months:
        st.subheader("지표 미리보기")
        selected = st.selectbox("연월 선택", metric_months, index=len(metric_months) - 1, key="metrics_month_select")
        rows = vibecoding.load_metric_rows(st.session_state.master_path, selected)
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

        st.subheader("그래프 미리보기")
        rate_img = os.path.join(graph_dir, f"집행률_비교_{selected}.png")
        delta_img = os.path.join(graph_dir, f"증감_비교_{selected}.png")
        c1, c2 = st.columns(2)
        if os.path.exists(rate_img):
            c1.image(rate_img, caption="집행률 비교")
        else:
            c1.caption("집행률 그래프 없음")
        if os.path.exists(delta_img):
            c2.image(delta_img, caption="전월 대비 증감 비교")
        else:
            c2.caption("전월 데이터가 없어 증감 그래프가 생성되지 않았습니다.")

    st.button("다음: 보고 자료 준비 →", type="primary", on_click=set_stage, args=("prepare",))


def page_prepare():
    st.header("📝 보고 자료 준비")
    st.caption("Step 8 — 검증·지표·그래프·담당자 회신을 하나의 자료로 통합하고, 보고서 목차 템플릿을 준비합니다.")

    graph_dir = derived_dir("그래프")
    email_dir = derived_dir("이메일")
    output_dir = derived_dir("공문자료")
    template_path = os.path.join(os.path.dirname(st.session_state.master_path) or ".", "월간예산집행보고서_템플릿.md")

    if st.button("보고 자료 생성", type="primary"):
        with st.spinner("보고 자료 통합 및 템플릿 준비 중..."):
            # run_final_doc_pipeline()은 검증로그_연도-월 시트가 실제 파일에 저장되어 있어야
            # 그 달을 대상으로 잡는다(find_months_needing_final_doc, 노트북 원본 로직 그대로).
            # 화면에는 검증 결과가 항상 최신으로 보이지만, 그건 매번 다시 계산해서 보여주는
            # 것일 뿐 파일에 저장된다는 보장이 없다 — '데이터 검증' 화면에서 '검증 실행'을
            # 실제로 누르지 않고 넘어온 경우 여기서 조용히 아무것도 안 만들어지는 문제가
            # 있었다. run_validation_pipeline()은 이미 검증된 달은 건너뛰므로, 여기서 먼저
            # 한 번 더 불러 빠진 달만 채워도 안전하다.
            _, v_log = run_captured(vibecoding.run_validation_pipeline, st.session_state.master_path)
            _, t_log = run_captured(vibecoding.create_report_template, template_path, vibecoding.TEMPLATE_CONTENT)
            _, d_log = run_captured(
                vibecoding.run_final_doc_pipeline, st.session_state.master_path, graph_dir, email_dir, output_dir
            )
        st.session_state.prepare_log_text = v_log + "\n" + t_log + "\n" + d_log
        st.session_state.stage_status["prepare"] = "done"
        st.rerun()

    if st.session_state.get("prepare_log_text"):
        st.subheader("실행 로그")
        st.code(st.session_state.prepare_log_text, language=None)

    if os.path.exists(template_path):
        with st.expander("📋 보고서 목차 템플릿 보기"):
            st.markdown(vibecoding.load_text(template_path))

    if os.path.isdir(output_dir):
        docs = sorted(f for f in os.listdir(output_dir) if f.endswith(".md"))
        if docs:
            st.subheader("생성된 공문작성자료")
            selected = st.selectbox("파일 선택", docs, index=len(docs) - 1, key="prepare_doc_select")

            # 공문작성자료 md는 그래프를 이미지 태그가 아니라 파일 경로 텍스트로만
            # 언급하도록 되어 있어(원본 노트북 설계), 여기서는 실제 그래프 파일을 찾아
            # 별도로 보여준다.
            sheet_name = selected.removeprefix("공문작성자료_").removesuffix(".md")
            rate_img = os.path.join(graph_dir, f"집행률_비교_{sheet_name}.png")
            delta_img = os.path.join(graph_dir, f"증감_비교_{sheet_name}.png")
            if os.path.exists(rate_img) or os.path.exists(delta_img):
                st.caption("첨부 그래프")
                c1, c2 = st.columns(2)
                if os.path.exists(rate_img):
                    c1.image(rate_img, caption="집행률 비교")
                if os.path.exists(delta_img):
                    c2.image(delta_img, caption="전월 대비 증감 비교")

            st.markdown(vibecoding.load_text(os.path.join(output_dir, selected)))

    st.button("다음: 보고서 초안 →", type="primary", on_click=set_stage, args=("draft",))


def page_draft():
    st.header("🤖 보고서 초안")
    st.caption("Step 9 — LLM(Upstage)으로 보고서 초안을 작성합니다.")
    st.info(
        "이 화면에서는 미리보기만 하고 디스크에 저장하지 않습니다 — 실제 저장은 "
        "'✅ 검토 및 저장' 화면에서 '최종 승인 및 저장'을 눌렀을 때 이루어집니다.",
        icon="ℹ️",
    )

    output_dir = derived_dir("공문자료")
    template_path = os.path.join(os.path.dirname(st.session_state.master_path) or ".", "월간예산집행보고서_템플릿.md")
    draft_dir = derived_dir("보고서초안")  # 이미 최종 승인·저장된 달을 확인하는 용도로만 읽는다

    if not os.path.exists(template_path) or not os.path.isdir(output_dir) or not os.listdir(output_dir):
        st.warning("먼저 '📝 보고 자료 준비' 화면에서 보고 자료를 생성해 주세요.", icon="⚠️")
        return

    if st.button("보고서 초안 생성 (Upstage LLM)", type="primary"):
        with st.spinner("LLM으로 보고서 초안 작성 중..."):
            log_lines = []
            template_text = vibecoding.load_text(template_path)
            for filename in sorted(os.listdir(output_dir)):
                match = vibecoding.DATA_PATTERN.match(filename)
                if not match:
                    continue
                sheet_name = match.group(1)
                final_saved = os.path.exists(os.path.join(draft_dir, f"보고서초안_{sheet_name}.md"))
                if final_saved:
                    log_lines.append(f"건너뜀 (이미 최종 저장됨): {sheet_name}")
                    continue
                data_text = vibecoding.load_text(os.path.join(output_dir, filename))
                log_lines.append(f"[{sheet_name}] 보고서 초안 생성 중...")
                content = vibecoding.draft_report(sheet_name, template_text, data_text)
                st.session_state.draft_previews[sheet_name] = content
                log_lines.append(f"미리보기 준비 완료: {sheet_name} (저장은 최종 승인 시 진행)")
        st.session_state.draft_log_text = "\n".join(log_lines)
        st.session_state.stage_status["draft"] = "done"
        st.rerun()

    if st.session_state.get("draft_log_text"):
        st.subheader("실행 로그")
        st.code(st.session_state.draft_log_text, language=None)

    saved_months = sorted(
        m.group(1) for f in (os.listdir(draft_dir) if os.path.isdir(draft_dir) else [])
        if (m := re.match(r"^보고서초안_(\d{4}-\d{2})\.md$", f))
    )
    preview_months = sorted(st.session_state.draft_previews)
    all_months = sorted(set(saved_months) | set(preview_months))

    if all_months:
        st.subheader("생성된 보고서 초안")
        labels = {m: (f"{m} (최종 저장됨)" if m in saved_months else f"{m} (미리보기 · 저장 전)") for m in all_months}
        selected = st.selectbox(
            "연월 선택", all_months, format_func=lambda m: labels[m], index=len(all_months) - 1, key="draft_month_select"
        )
        if selected in st.session_state.draft_previews:
            content = st.session_state.draft_previews[selected]
        else:
            content = vibecoding.load_text(os.path.join(draft_dir, f"보고서초안_{selected}.md"))
        base_dir = os.path.dirname(st.session_state.master_path) or "."
        render_markdown_with_local_images(content, base_dir)

    st.button("다음: 검토 및 저장 →", type="primary", on_click=set_stage, args=("review",))


def page_review():
    st.header("✅ 검토 및 저장")
    st.caption("Step 10-11 — 사람이 초안을 검토하고 승인 후 최종 저장합니다.")

    draft_dir = derived_dir("보고서초안")

    saved_months = sorted(
        m.group(1) for f in (os.listdir(draft_dir) if os.path.isdir(draft_dir) else [])
        if (m := re.match(r"^보고서초안_(\d{4}-\d{2})\.md$", f))
    )
    preview_months = sorted(st.session_state.draft_previews)
    all_months = sorted(set(saved_months) | set(preview_months))

    if not all_months:
        st.warning("먼저 '🤖 보고서 초안' 화면에서 초안을 생성해 주세요.", icon="⚠️")
        return

    labels = {m: (f"{m} (최종 저장됨)" if m in saved_months else f"{m} (미리보기 · 저장 전)") for m in all_months}
    selected = st.selectbox(
        "검토할 연월 선택", all_months, format_func=lambda m: labels[m], index=len(all_months) - 1, key="review_select"
    )
    already_saved = selected not in st.session_state.draft_previews
    if already_saved:
        content = vibecoding.load_text(os.path.join(draft_dir, f"보고서초안_{selected}.md"))
    else:
        content = st.session_state.draft_previews[selected]

    base_dir = os.path.dirname(st.session_state.master_path) or "."
    render_markdown_with_local_images(content, base_dir)

    st.subheader("검토 체크리스트")
    c1 = st.checkbox("표의 숫자가 공문작성자료 원자료와 일치하는가", key="chk_numbers")
    c2 = st.checkbox("LLM이 만든 문장 중 원자료에 없는 내용(특히 비교 수치)을 지어낸 부분은 없는가", key="chk_fabrication")
    c3 = st.checkbox("이상치·누락값에 대한 사유가 담당자 회신 내용과 정확히 일치하는가", key="chk_reason")
    c4 = st.checkbox("경영진 확인·결정 요청 사항이 실제로 결재가 필요한 내용인가(과장되지 않았는가)", key="chk_decision")
    c5 = st.checkbox("오탈자·문장 어색함 등 표현상 문제가 없는가", key="chk_wording")

    all_checked = all([c1, c2, c3, c4, c5])
    if already_saved:
        st.caption(f"이미 최종 저장된 초안입니다. 수정이 필요하면 `보고서초안_{selected}.md` 파일을 직접 편집하세요.")
    if st.button("최종 승인 및 저장", type="primary", disabled=not all_checked):
        os.makedirs(draft_dir, exist_ok=True)
        draft_path = os.path.abspath(os.path.join(draft_dir, f"보고서초안_{selected}.md"))
        with open(draft_path, "w", encoding="utf-8") as f:
            f.write(content)

        st.session_state.draft_previews.pop(selected, None)
        st.session_state.stage_status["review"] = "done"
        st.session_state.review_saved_path = draft_path
        st.rerun()
    if not all_checked:
        st.caption("모든 체크리스트 항목을 확인해야 최종 저장할 수 있습니다.")

    if st.session_state.get("review_saved_path"):
        st.success("최종 보고서가 저장되었습니다. 아래 위치를 확인해 주세요.", icon="✅")
        st.code(st.session_state.review_saved_path, language=None)
        st.caption(
            "조직의 전자결재 시스템에 이 파일 내용을 옮겨 결재를 상신하세요. "
            "다음 달에는 새 부서 제출 파일로 다시 Step 1부터 시작하면 됩니다."
        )


PAGES = {
    "collect": page_collect,
    "validate": page_validate,
    "issue": page_issue,
    "revalidate": page_revalidate,
    "metrics": page_metrics,
    "prepare": page_prepare,
    "draft": page_draft,
    "review": page_review,
}


def main():
    init_state()
    render_sidebar()
    PAGES[st.session_state.current_stage]()


if __name__ == "__main__":
    main()
