# =========================================================================
# 2-BO'LIM: RANGLI YIG'MA SHABLON HISOBOT (TO'G'RILANGAN VERSIYA)
# =========================================================================
with tab2:
    st.markdown("""
    <div class="portal-banner" style="background: linear-gradient(135deg, #065f46 0%, #047857 50%, #059669 100%);">
        <h2>📑 Rangli Yig'ma Hisobot Shablonini Yangilash</h2>
        <p>Birlamchi ma'lumotlarni hisoblab, rangli Excel shablonning "Kunlik" va "Tadbir boshidan" ustunlariga avtomatik integratsiya qilish</p>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("##### 📌 1-qadam: Sug'urta turini tanlang")
    report_type_selection = st.radio(
        "Sug'urta yo'nalishi:",
        ["Kuzatuv (Kafolat) — 75% / 25%", "OSAGO (Majburiy) — 80% / 20%"],
        horizontal=True,
        key="tab2_report_type"
    )
    is_kuzatuv = "Kuzatuv" in report_type_selection

    st.markdown("---")
    st.markdown("##### 📁 2-qadam: Ikkala Excel faylni yuklang")
    f_col1, f_col2 = st.columns(2)
    with f_col1:
        daily_file = st.file_uploader(
            "1. Birlamchi kunlik Excel fayl (Polislar):",
            type=["xlsx", "xls"],
            key="tab2_daily_file"
        )
    with f_col2:
        template_file = st.file_uploader(
            "2. Rangli shablon Excel fayl (Namunadagi jadval):",
            type=["xlsx", "xls"],
            key="tab2_template_file"
        )

    if daily_file and template_file:
        if st.button("⚡ Hisobotni yangilash va rangli Excel yaratish", type="primary", use_container_width=True):
            with st.spinner("Hisobot to'g'ri ustunlar bo'yicha to'ldirilmoqda..."):
                try:
                    df_in = pd.read_excel(daily_file)

                    # 1. Birlamchi fayl ustunlarini aniqlash
                    cols_lower = {str(c).lower().strip(): c for c in df_in.columns}
                    f_col = None
                    for cand in ["sug'urta firma", "sugurta_firma", "insurance_company", "firma"]:
                        if cand in cols_lower:
                            f_col = cols_lower[cand]
                            break
                    if not f_col:
                        f_col = df_in.columns[0]

                    s_col = None
                    for cand in ["100%", "res_ins_prem", "ins_prem", "premium", "mukofot_sum"]:
                        if cand in cols_lower:
                            s_col = cols_lower[cand]
                            break
                    if not s_col:
                        s_col = df_in.columns[2]

                    pct_org = 0.75 if is_kuzatuv else 0.80
                    pct_mbl = 0.25 if is_kuzatuv else 0.20

                    # JAMI qatorini hisobdan chiqarish
                    df_in = df_in[~df_in[f_col].astype(str).str.upper().str.contains("JAMI|ИТОГ|TOTAL")].copy()
                    df_in[s_col] = pd.to_numeric(
                        df_in[s_col].astype(str).str.replace(" ", "").str.replace(",", "."),
                        errors="coerce"
                    ).fillna(0)

                    # Firma nomlarini tozalash va solishtirish funksiyasi
                    def clean_firm(val):
                        v = str(val).upper()
                        v = re.sub(r'["“”`\'\s]', '', v)
                        for token in ['AJ', 'SK', 'EISK', 'QK', 'KOMPANIYASI', 'SUGURTA', 'SUG`URTA', 'INSURANCE']:
                            v = v.replace(token, '')
                        return v

                    df_in["clean_firm"] = df_in[f_col].apply(clean_firm)

                    # Firmalar bo'yicha yig'ish
                    grouped = df_in.groupby("clean_firm").agg(
                        cnt=(s_col, "count"),
                        sum_val=(s_col, "sum")
                    ).to_dict(orient="index")

                    # Jami kunlik summa (mln so'mda)
                    total_daily_sum_mln = sum(data["sum_val"] for data in grouped.values()) / 1_000_000.0

                    # 2. Rangli shablonni ochish
                    wb = openpyxl.load_workbook(template_file)
                    ws = wb.active

                    def to_float(val):
                        if val is None:
                            return 0.0
                        if isinstance(val, (int, float)):
                            return float(val)
                        s = str(val).replace(" ", "").replace(",", ".")
                        try:
                            return float(s)
                        except:
                            return 0.0

                    # Ustunlar koordinatalari:
                    # F: Tadbir polislar, G: Tadbir mukofot, H: Tadbir 75%, I: Tadbir 25%
                    # J: Kunlik polislar, K: Kunlik mukofot, L: Kunlik 75%, M: Kunlik 25%, N: foiz ulushi
                    COL_T_CNT = 6
                    COL_T_PREM = 7
                    COL_T_ORG = 8
                    COL_T_MBL = 9

                    COL_K_CNT = 10
                    COL_K_PREM = 11
                    COL_K_ORG = 12
                    COL_K_MBL = 13
                    COL_K_PCT = 14

                    # Asosiy firmalar qatorlari (10-qatordan 22-qatorgacha)
                    total_k_cnt = 0
                    total_k_prem = 0.0
                    total_k_org = 0.0
                    total_k_mbl = 0.0

                    for r in range(10, 23):
                        firm_name_raw = ws.cell(r, 3).value  # C ustuni
                        if not firm_name_raw:
                            continue

                        c_name = clean_firm(firm_name_raw)
                        matched_key = next((k for k in grouped if k in c_name or c_name in k), None)

                        if matched_key:
                            d_cnt = grouped[matched_key]["cnt"]
                            d_prem = grouped[matched_key]["sum_val"] / 1_000_000.0
                        else:
                            d_cnt = 0
                            d_prem = 0.0

                        d_org = d_prem * pct_org
                        d_mbl = d_prem * pct_mbl

                        total_k_cnt += d_cnt
                        total_k_prem += d_prem
                        total_k_org += d_org
                        total_k_mbl += d_mbl

                        # Kunlik ustunlarini to'ldirish
                        ws.cell(r, COL_K_CNT, value=d_cnt if d_cnt > 0 else "-")
                        ws.cell(r, COL_K_PREM, value=round(d_prem, 4) if d_prem > 0 else "-")
                        ws.cell(r, COL_K_ORG, value=round(d_org, 4) if d_org > 0 else "-")
                        ws.cell(r, COL_K_MBL, value=round(d_mbl, 4) if d_mbl > 0 else "-")

                        # Foiz ulushi (N ustuni)
                        if total_daily_sum_mln > 0 and d_prem > 0:
                            ws.cell(r, COL_K_PCT, value=f"=K{r}*100/K$9")
                        else:
                            ws.cell(r, COL_K_PCT, value="-" if d_prem == 0 else 0)

                        # Tadbir boshidan ustunlariga qo'shish (kümülyativ)
                        if d_cnt > 0:
                            cur_t_cnt = to_float(ws.cell(r, COL_T_CNT).value)
                            cur_t_prem = to_float(ws.cell(r, COL_T_PREM).value)
                            cur_t_org = to_float(ws.cell(r, COL_T_ORG).value)
                            cur_t_mbl = to_float(ws.cell(r, COL_T_MBL).value)

                            ws.cell(r, COL_T_CNT, value=int(cur_t_cnt + d_cnt))
                            ws.cell(r, COL_T_PREM, value=round(cur_t_prem + d_prem, 4))
                            ws.cell(r, COL_T_ORG, value=round(cur_t_org + d_org, 4))
                            ws.cell(r, COL_T_MBL, value=round(cur_t_mbl + d_mbl, 4))

                    # 3. JAMI qatorini (9-qator) formulalar bilan tiklash
                    ws.cell(9, COL_T_CNT, value="=SUM(F10:F22)")
                    ws.cell(9, COL_T_PREM, value="=SUM(G10:G22)")
                    ws.cell(9, COL_T_ORG, value="=SUM(H10:H22)")
                    ws.cell(9, COL_T_MBL, value="=SUM(I10:I22)")

                    ws.cell(9, COL_K_CNT, value="=+J10+J11+J12+J13+J14+J15+J16+J17+J18")
                    ws.cell(9, COL_K_PREM, value="=+K10+K11+K12+K13+K14+K15+K16+K17+K18")
                    ws.cell(9, COL_K_ORG, value="=+L10+L11+L12+L13+L14+L15+L16+L17+L18")
                    ws.cell(9, COL_K_MBL, value="=+M10+M11+M12+M13+M14+M15+M16+M17+M18")
                    ws.cell(9, COL_K_PCT, value=1)  # 100%

                    # 4. Pastdagi Oktabr oyi qatoriga (39-qator) kunlik ko'rsatkichlarni qo'shish
                    if ws.cell(39, 2).value and "Oktabr" in str(ws.cell(39, 2).value):
                        c_cnt_39 = to_float(ws.cell(39, 4).value)
                        c_prem_39 = to_float(ws.cell(39, 6).value)
                        c_org_39 = to_float(ws.cell(39, 9).value)
                        c_mbl_39 = to_float(ws.cell(39, 12).value)

                        ws.cell(39, 4, value=int(c_cnt_39 + total_k_cnt))
                        ws.cell(39, 6, value=round(c_prem_39 + total_k_prem, 4))
                        ws.cell(39, 9, value=round(c_org_39 + total_k_org, 4))
                        ws.cell(39, 12, value=round(c_mbl_39 + total_k_mbl, 4))

                    out_stream = io.BytesIO()
                    wb.save(out_stream)
                    st.session_state.filled_template_bytes = out_stream.getvalue()

                    st.success("✅ Shablon to'liq va aniq hisob-kitoblar bilan yangilandi!")

                except Exception as ex:
                    st.error(f"Xatolik yuz berdi: {ex}")

    if st.session_state.filled_template_bytes is not None:
        st.markdown("---")
        st.download_button(
            label="📥 Yangilangan Rangli Hisobotni Yuklab Olish (.xlsx)",
            data=st.session_state.filled_template_bytes,
            file_name="yangilangan_kuzatuv_hisoboti.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
