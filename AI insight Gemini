# ============================================================
# GEMINI VERSION — free-tier, no billing account required.
#
# WHERE THIS GOES — same 3 insertion points as before:
#
# 1) Near your other imports at the top of the file, add:
#
#       from google import genai
#
#    Install it once with:  pip install google-genai
#    (this replaces the earlier `anthropic` package — you don't
#    need both)
#
# 2) Anywhere before the TABS section (e.g. right after your
#    EXCEL/PDF EXPORT functions), paste build_ratio_prompt() and
#    get_ai_interpretation() below. build_ratio_prompt() is IDENTICAL
#    to the Claude version — only get_ai_interpretation() changes,
#    since that's the only part that talks to a specific vendor's API.
#
# 3) Inside `with tab_summary:`, AFTER your existing dl1/dl2/dl3
#    download-button block and BEFORE the final st.caption(...)
#    reminder, paste the UI BLOCK at the bottom of this file.
#
# Also add the SIDEBAR ADDITION below inside your `with st.sidebar:`
# block, so there's a place to enter the free API key.
#
# Getting a free key: go to aistudio.google.com/apikey, sign in with
# a Google account, click "Create API key." No credit card needed for
# the free tier. Model used below (gemini-2.5-flash) is on Google's
# free-tier-eligible list as of Aug 2026 — if you're setting this up
# later and it's been a while, check aistudio.google.com/models for
# the current recommended free Flash model name, since Google renames/
# releases new versions fairly often.
# ============================================================


# ============================================================
# 2) FUNCTIONS — paste before the TABS section
# ============================================================

def build_ratio_prompt(companies, benchmark_sector, matrix_rows):
    """Turn the already-computed ratio matrix into a compact text prompt.
    Sends COMPUTED RATIOS only (not raw financial statement line items) —
    smaller prompt, and there's nothing sensitive in public-company ratios
    anyway, but it's a good habit for when this pattern gets reused on
    non-public data later. (Same function as the Claude version — the
    prompt-building logic doesn't depend on which AI vendor answers it.)"""
    lines = [
        f"You are a finance teaching assistant. Companies being compared: "
        f"{', '.join(c['ticker'] for c in companies)}. "
        f"Industry benchmark used: {benchmark_sector}.",
        "",
        "Ratio data (Year1 -> Year2, vs benchmark):",
    ]
    for row in matrix_rows:
        vals = []
        ci = 0
        for c in companies:
            v1 = row["company_values"][ci]
            v2 = row["company_values"][ci + 1]
            vals.append(f"{c['ticker']}: {fmt(v1, row['suffix'])} -> {fmt(v2, row['suffix'])}")
            ci += 2
        bench = fmt(row["benchmark"], row["suffix"])
        lines.append(f"- {row['ratio']} ({row['category']}): " + "; ".join(vals) + f" | benchmark: {bench}")
    lines += [
        "",
        "Write a short (250-350 word) plain-English interpretation for finance "
        "students: (1) each company's key strengths/weaknesses vs. the benchmark, "
        "(2) one notable year-over-year trend per company, (3) one thing a student "
        "should double-check or be skeptical of in this data. Do not invent numbers "
        "beyond what's given above.",
    ]
    return "\n".join(lines)


def get_ai_interpretation(prompt_text, api_key, model="gemini-2.5-flash"):
    """Calls the Gemini API and returns (text, error). Never raises —
    Streamlit apps should degrade gracefully in front of a classroom."""
    if not api_key:
        return None, "No API key provided. Add a free Gemini key in the sidebar to use this feature."
    try:
        client = genai.Client(api_key=api_key)
        response = client.models.generate_content(
            model=model,
            contents=prompt_text,
        )
        return response.text, None
    except Exception as e:
        return None, f"AI request failed: {e}"


# ============================================================
# SIDEBAR ADDITION — one block inside your existing `with st.sidebar:`
# (put it near the top, before "How to use this tool" or after it)
# ============================================================
#
#     st.divider()
#     st.markdown("**AI Interpretation (optional, free tier)**")
#     ai_api_key = st.text_input(
#         "Gemini API key", type="password",
#         value=os.environ.get("GEMINI_API_KEY", ""),
#         help="Free — no credit card needed. Get one at "
#              "aistudio.google.com/apikey. For a whole class, set this "
#              "as a Streamlit Cloud secret instead of typing it per-session.",
#     )
#
# (needs `import os` at the top of the file if you don't already have it)


# ============================================================
# 3) UI BLOCK — paste inside `with tab_summary:`, after the
#    dl1/dl2/dl3 download-button block, before the final st.caption(...)
# ============================================================
#
#     st.divider()
#     st.subheader("🤖 AI-Generated Interpretation (Beta)")
#     st.caption(
#         "Generates a narrative read of the ratio table above using Gemini "
#         "(free tier). Treat this the same way you'd treat a classmate's "
#         "first draft: useful starting point, not a verified answer."
#     )
#
#     current_signature = (
#         tuple(c["ticker"] for c in companies),
#         tuple((c["y1_label"], c["y2_label"]) for c in companies),
#         benchmark_sector,
#     )
#
#     if st.button("Generate AI Analysis", key="ai_generate_btn"):
#         prompt_text = build_ratio_prompt(companies, benchmark_sector, matrix_rows)
#         with st.spinner("Asking Gemini..."):
#             text, err = get_ai_interpretation(prompt_text, ai_api_key)
#         if err:
#             st.error(err)
#         else:
#             st.session_state["ai_analysis_text"] = text
#             st.session_state["ai_prompt_text"] = prompt_text
#             st.session_state["ai_analysis_signature"] = current_signature
#
#     if "ai_analysis_text" in st.session_state:
#         if st.session_state.get("ai_analysis_signature") != current_signature:
#             st.warning(
#                 "⚠️ The company/year/benchmark selection has changed since this "
#                 "analysis was generated — click Generate again for it to match "
#                 "what's currently shown above."
#             )
#         st.info(st.session_state["ai_analysis_text"])
#
#         with st.expander("📋 Exact prompt sent to the AI (for your disclosure statement)"):
#             st.code(st.session_state["ai_prompt_text"], language="text")
#
#         st.markdown("**Your verification** — required before you use this in an assignment:")
#         st.text_area(
#             "Does this match your own read of the numbers? What, if anything, "
#             "would you correct or add?",
#             key="student_ai_critique", height=150,
#         )
