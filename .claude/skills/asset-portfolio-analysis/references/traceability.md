# مصفوفة التتبع: كل قاعدة في الموجّه الأصلي وموضع تطبيقها

> **الغرض:** إثبات أن كل جملة أو قاعدة في `00-original-brief.md` لها موضع تطبيق واحد أساسي، وطريقة إنفاذ.
> **الترقيم:** رقم القسم في الموجّه، ثم رقم القاعدة داخله بترتيب ورودها. الوسم نفسه، مثل [§6.18]، يظهر في الملف الأساسي بجوار القاعدة.
> **حجية النص الأصلي:** عند أي تعارض أو غموض يُرجع إلى النص الأصلي، كما تنص مقدمته.
> **العدد الكلي: 156 قاعدة** موزعة كما يلي:

| القسم | 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9 | 10 | 11 | 12 | 13 |
|---|---|---|---|---|---|---|---|---|---|---|---|---|---|
| عدد القواعد | 6 | 6 | 13 | 12 | 11 | 34 | 12 | 4 | 8 | 8 | 22 | 7 | 13 |

## رموز الملفات والإنفاذ

| رمز الملف | الملف |
|---|---|
| SK | `SKILL.md` |
| SC | `references/scope-and-classification.md` |
| DC | `references/data-cleaning.md` |
| RF | `references/reference-file-controls.md` |
| MT | `references/metrics.md` |
| DA | `references/diagnosis-and-assessment.md` |
| RS | `references/recommendations-and-scenarios.md` |
| DS | `references/dashboard-spec.md` |
| UV | `references/updates-and-versioning.md` |
| DL | `references/delivery-checklist.md` |

| رمز الإنفاذ | الأداة |
|---|---|
| `inspect` | `inspect_workbook.py`، ونوع الشذوذ بعده |
| `normalize` | `normalize.py`، ونوع المشكلة بعده |
| `compute` | `compute_metrics.py`، ومعه المحرك |
| `V:` | فحص في `validate_analysis.py`، ويُذكر «مانع» إن كان يمنع البناء |
| `C:` | فحص في `check_dashboard.mjs` |
| `build` | `build_dashboard.py` والقالب |
| `export` | `export_workbook.py` |
| `compare` | `compare_versions.py` |
| `schema` | مخططات `schemas/` |
| «حكم المحلل» | لا فحص آلي، والالتزام مسؤولية المحلل |

## 1. الدور والهدف

| # | القاعدة | الملف | الإنفاذ |
|---|---|---|---|
| 1.1 | الدور: محلل محفظة أصول ومستشار تطوير استثمارات ومحلل بيانات ومصمم لوحات للإدارة والمجلس | SK | حكم المحلل |
| 1.2 | عند إرفاق Excel: دراسة متعمقة ثم داشبورد عربي واضح وسهل الاستخدام | DL §1 | `build`؛ `C:` كل الفحوص |
| 1.3 | الأهداف: دخل مستدام، محافظة على الأصول، كفاءة الاستخدام، استمرارية خدمات الأيتام | RS §1 | حكم المحلل (`rationale`) |
| 1.4 | العمق في التفاصيل، والواجهة التنفيذية مختصرة | DS §1 | `build` (أقسام التفاصيل)؛ `C:structure`، `C:details_sections` |
| 1.5 | الإجابة عن الأسئلة الخمسة | DA §1 | `V:exec_summary_elements` (جزئي)؛ حكم المحلل |
| 1.6 | الأصل الخدمي غرضه مختلف، ولا يُحكم بضعفه لعدم الإيجار | DA §1 | بُعد `use` ونطاق العائد الاستثماري (`compute`)؛ `build` (لا عائد للخدمي في الجدول) |

## 2. طريقة التشغيل

| # | القاعدة | الملف | الإنفاذ |
|---|---|---|---|
| 2.1 | ابدأ بالفحص واستمر حتى المخرجات، ولا تكتفِ بخطة أو وصف | DL §1 | تسلسل DL §4؛ `build`؛ `C:` |
| 2.2 | الجهة والفترة وتاريخ الدراسة والعملة وتواريخ التقييم، ولا افتراض أن سنة التحليل هي سنة البيانات | SC §1 | حقول السياق في `02`؛ `inspect` (`title_cells`، `years_in_header`) |
| 2.3 | حفظ الأصل، مخرجات مستقلة، لا نشر ولا إرسال | DL §1 | فتح للقراءة فقط؛ `V:source_hash`؛ حكم المحلل (لا Artifact) |
| 2.4 | أنجز الممكن، وبيّن أثر النقص، ولا يعطل نقص مؤشر الدراسة | DL §1 | `compute` (حالة وسبب لكل مؤشر مستقل)؛ `data_requests` |
| 2.5 | ثلاثة أسئلة مؤثرة كحد أقصى في كل دفعة | DL §1 | `V:questions_max` (مانع)؛ `schema` (`maxItems 3`) |
| 2.6 | ميّز المعلومة والحساب والاستنتاج والافتراض، ولا تختلق | DA §2 | `kind` و`evidence`؛ `V:gap_evidence` (مانع)، `V:metric_tokens` (مانع)، `V:evidence_refs_other`، `V:hand_typed_numbers` |

## 3. فهم الملف وتحديد النطاق

| # | القاعدة | الملف | الإنفاذ |
|---|---|---|---|
| 3.1 | فحص كل الأوراق (المخفية والجداول والصيغ والملاحظات) وخريطتها | SC §2 | `inspect` (`hidden_sheet`، الجداول، الصيغ، التعليقات)؛ `sheets[]` في `02`؛ `build` (قسم `workbook`) |
| 3.2 | لا جمع تلقائي للأوراق، وتحقق انتمائها | SC §3 | `decision` و`default_scope: isolated`؛ `base_where` في `compute` |
| 3.3 | مطابقة الأصول بالمعرّفات والخصائص قبل استخدام إيراداتها ومصروفاتها | SC §3 | `links[]` (المؤكد وحده)؛ `inspect:cross_sheet_name_match` تلميح |
| 3.4 | عزل غير المتحقق من نطاقه مع السبب | SC §3 | `scope_reason` و`exclusions`؛ `V:isolated_in_metrics`؛ `export` («المستبعدات») |
| 3.5 | تحديد ما يمثله الصف | SC §4 | `record_type`؛ `inspect` (`row_kinds`)؛ `normalize:total_like_row` |
| 3.6 | لا جمع مزدوج للأصل ووحداته ولا للمجاميع وتفاصيلها | SC §4 | `exclude_rows`؛ `V:total_rows_in_scope` (مانع)؛ `V:duplicate_record_ids` (مانع) |
| 3.7 | الملكية والحقوق غير حصرية، والوقف مستقل | SC §5 | `schema` (`ownership` قائمة، `waqf` مستقل)؛ `compute` (توزيع «قد تتداخل الفئات») |
| 3.8 | الاستخدام بفئاته الخمس | SC §5 | `vocab`؛ `V:unknown_classification`؛ `normalize:missing_classification` |
| 3.9 | طبيعة الدخل بفئاتها | SC §5 | `vocab`؛ `V:unknown_classification` |
| 3.10 | النوع والموقع والحالة وتخصيص الريع والقيود المثبتة | SC §5 | أعمدة `dim`؛ `breakdowns`؛ `filters` |
| 3.11 | إيرادات المدارس والأنشطة ليست إيجارًا | SC §6 | `operating_activity` و`operating`؛ `compute` (`activity_revenue`، واستبعاد التشغيلي من العائد) |
| 3.12 | إيراد العقار المدار للغير ليس دخلًا للجمعية | SC §6 | `managed_for_others`؛ `compute` (مستبعد من كل المؤشرات القياسية) |
| 3.13 | الأصل كاملًا أم الحصة، قبل نسب الملكية | SC §7 | `partial` و`derived_fields` صريحة؛ `assumptions`؛ حكم المحلل |

## 4. تنظيف البيانات

| # | القاعدة | الملف | الإنفاذ |
|---|---|---|---|
| 4.1 | معرّف ثابت ومصدر (ملف، ورقة، صف أو خلية) | DC §1 | `normalize` (`id_rule`، `ref`، `source`)؛ `V:duplicate_record_ids` (مانع) |
| 4.2 | توحيد الأرقام والتواريخ والعملات والمساحات والمسميات مع حفظ الأصل | DC §2 | `parse_number` و`normalize_arabic`؛ `raw` |
| 4.3 | التحقق من التقويم قبل تحويل التواريخ | DC §3 | `calendar`؛ `parse_date` (أم القرى)؛ `inspect:hijri_date_text`؛ `normalize:hijri_conversion` و`date_unparsed` |
| 4.4 | فصل الصفر المؤكد والفارغ وغير المتاح وغير المنطبق، ولا صفر للمجهول | DC §4 | حالات TV و`zero_policy`؛ `V:tv_status_consistency` (مانع) |
| 4.5 | لا حذف للأسماء المتشابهة، ومطابقة المعرّف والصك والموقع والمساحة، والصك المشترك | DC §5 | `inspect:duplicate_name` و`shared_deed`؛ `normalize:possible_duplicate_name` و`shared_deed` |
| 4.6 | محتوى الحقل لا عنوانه: «غير مقيد» ليست مبالغ، والأدوار ليست وحدات | DC §6 | `kind: text` و`count_mixed`؛ `inspect:text_in_numeric` و`mixed_types`؛ `normalize:count_mixed_text` |
| 4.7 | مراجعة الصيغ وقيمها المخزنة وأخطاء المراجع | DC §7 | `inspect:formula_year_mismatch`، `cached_value_mismatch`، `cached_missing`، `excel_error`؛ `normalize:formula_no_cache` |
| 4.8 | إعادة حساب المجاميع ومطابقتها مع المدونة وإظهار الفروق | DC §7 | `compute` (`sum`)؛ `inspect:total_mismatch`، `hardcoded_total`، `count_mismatch`؛ `normalize:valuation_total_mismatch` |
| 4.9 | فصل الضريبة بحسب البيانات، ولا نسبة واحدة، ولا تعديل بالاسم | DC §8 | `vat` في `field_meta`؛ `compute` (اختيار `excl`)؛ `inspect:vat_identical`؛ `normalize:vat_incl_lt_excl` |
| 4.10 | فعلي أم متوقع، مستحق أم محصل، سنوي أم جزئي، ولا وصف غير مثبت | DC §9 | `basis`، `cash_basis`، `completeness`؛ `compute` (قابلية المقارنة)؛ `inspect:partial_period_header` |
| 4.11 | سجل مشكلات البيانات (المشكلة، الموضع، الأثر، المعالجة) | DC §10 | `normalize` (`DQ-xxx`)؛ `data_issues_extra`؛ `export`؛ `build` (قسم `issues`) |
| 4.12 | فصل جودة البيانات عن أداء المحفظة | DC §10 | `schema` (بُعد `data_quality` مستقل)؛ حكم المحلل |

## 5. ضوابط الملف المرجعي

| # | القاعدة | الملف | الإنفاذ |
|---|---|---|---|
| 5.0 | التطبيق عند مطابقة البنية بعد إعادة التحقق، دون تثبيت الصفوف أو الأرقام | RF §1 | فحص البصمة؛ `inspect` لكل إصدار؛ `methodology_notes` |
| 5.1 | سنوات الإيراد والتقييم، وأحدث تقييم ليس قيمة سوقية حالية | RF §2 | `field_meta`؛ `compute` (`valuation.note`) |
| 5.2 | فصل المقيّم عن المجمع والمدارس، ولا إيراد في البسط بلا مقام | RF §2 | `leased_land_project` و`operating_activity`؛ النسبة الزوجية؛ `V:pairwise_sets` (مانع) |
| 5.3 | حدود إيراد المجمع والمدارس، والإيجارات الداخلية، ولا ازدواج بلا دليل | RF §2 | `field_overrides` و`where`؛ `data_issues_extra`؛ حكم المحلل |
| 5.4 | أجرة الأرض بند تكلفة محتمل، وطلب بيان مستقل | RF §2 | لا ربط بـ`exp.*`؛ `data_requests` |
| 5.5 | فحص الورقتين المخفيتين، و«المنسدلات» ليست سجل أصول | RF §2 | `decision` (`isolated`/`excluded`)؛ `links` المؤكدة |
| 5.6 | «صافي الأصول» نصي، فلا يكون مقامًا | RF §2 | `kind: text` (`data-contract` §3)؛ `inspect:text_in_numeric` |
| 5.7 | «عدد الوحدات التأجيرية» مختلط، فلا جمع ولا إشغال منه | RF §2 | `count_mixed`؛ `normalize:count_mixed_text` |
| 5.8 | التحقق من إجمالي 2024 وأساس الضريبة والاكتمال قبل المقارنة | RF §2 | `compute` (`sum`، `comparable_with_prev`)؛ `V:growth_comparability` |
| 5.9 | الأسماء المكررة والتسلسل والصكوك، ولا عدد من آخر مسلسل، ولا تصرف في وقف من «غير مقيد» | RF §2 | `inspect:duplicate_name` و`serial_gap`؛ `asset_count`؛ بُعد `waqf` |
| 5.10 | المؤشرات المتأثرة بالغياب «غير متاحة» مع الاستمرار | RF §2 | `compute` (`noi` غير قابل للاحتساب)؛ `data_requests` |

## 6. المؤشرات المالية

| # | القاعدة | الملف | الإنفاذ |
|---|---|---|---|
| 6.1 | لكل مؤشر: تعريف، فترة، وحدة، نطاق، مصدر، أساس القيمة وتاريخها، حدود الاكتمال | MT §0 | `schema` (`MetricSpec`)؛ الحالة والمستبعد في النتيجة؛ `build` (النطاق والأساس) |
| 6.2 | مجموع القيم المثبتة مرة لكل أصل مع تاريخ التقييم وأساسه | MT §2 | `recorded_value`؛ `valuation` |
| 6.3 | لا جمع لتقييمات سنوات مختلفة، ولا للأرض والمبنى مع إجماليهما | MT §2 | `V:valuation_years_mixed`؛ `normalize:valuation_total_mismatch` |
| 6.4 | لا خلط للدفتري والتكلفة مع السوقي دون فصل وبيان الأثر | MT §2 | `valuation_basis`؛ حكم المحلل |
| 6.5 | تواريخ التقييم المختلفة يُصرَّح بها، ولا تقييم موحد | MT §2 | `date_text`؛ `V:valuation_years_mixed` |
| 6.6 | فصل غير المقيم والمدار للغير وحقوق الانتفاع والمشروعات التشغيلية | MT §2 | `excluded[]`؛ شروط `where` |
| 6.7 | «القيمة المسجلة للأصول المقيمة» ولا إيحاء بالقيمة الكاملة | MT §2 | التسمية القياسية؛ حكم المحلل |
| 6.8 | توزيع القيمة والإيراد حسب الأبعاد عند كفاية البيانات | MT §2 | `compute` (`breakdowns`)؛ `C:charts` |
| 6.9 | إجمالي الإيراد غير المكرر: الإيجارات منفصلة عن الأنشطة، والمتكرر عن الاستثنائي | MT §3 | `property_revenue` و`activity_revenue`؛ عزل المصدر المكرر؛ حقل `one_off` |
| 6.10 | تعريف عائد الدخل الإجمالي | MT §4.1 | `gross_yield_*`؛ اختيار حقل `excl` |
| 6.11 | صافي الدخل التشغيلي، وفصل الإهلاك والتمويل والرأسمالي، وأساس توزيع المشترك | MT §4.3 | `noi`؛ `exp_field`؛ `field_meta.note` |
| 6.12 | عائد الدخل التشغيلي الصافي | MT §4.3 | حقل `noi.<P>` المشتق ثم نسبة |
| 6.13 | العائد من مجاميع متطابقة، لا متوسط بسيط | MT §4.1 | المحرك (Σ/Σ)؛ `V:pairwise_sets` (مانع)؛ `V:metrics_reproducible` (مانع) |
| 6.14 | ثلاثة نطاقات منفصلة | MT §4.1 | `gross_yield_owned` و`investment` و`income_producing` |
| 6.15 | لا استبعاد لاستثماري إيراده صفري لتحسين العائد | MT §4.1 | الصفر حالة معروفة في المحرك؛ `V:yield_zero_revenue_scope` (مانع) |
| 6.16 | النسبة العامة لا تكفي للحكم على الخدمي | MT §4.1 | تعريف `gross_yield_owned` |
| 6.17 | «عائد الجزء القابل للاحتساب» مع المشمول والمستبعد وحصة الإيراد، ولا تغطية بالقيمة | MT §4.2 | حالة `partial`؛ `n_included` و`coverage_num_share`؛ المحرك لا يحسب تغطية القيمة؛ `V:partial_unnamed` |
| 6.18 | التسمية الدقيقة «الإيراد المسجل لعام P إلى تقييم Y…»، ولا صافي ولا كلي | MT §4.2 | `V:label_before_expenses`، `V:label_net`، `V:label_total_return`، `V:label_vat_basis`، `V:label_years`، `V:label_roi`، `V:narrative_metric_naming` (مانعة)؛ `V:narrative_claims`، `V:yield_vat_inclusive` |
| 6.19 | غياب المصروفات ليس صفرًا | MT §4.3 | `noi` غير قابل للاحتساب بسببه؛ `V:label_net` |
| 6.20 | فائض النشاط منفصل، ولا ربح من الإيراد وحده | MT §4.3 | استبعاد التشغيلي من العائد؛ تعريف `activity_revenue` |
| 6.21 | فترتان متماثلتان بأساس موحد، والتغير بالريال والنسبة | MT §5 | `compute` (قابلية المقارنة من `field_meta`، رفض الإعداد المخالف، `abs_change`، `pct`)؛ `V:growth_comparability_basis` و`V:period_fields` (مانعان)؛ `V:growth_comparability`، `V:growth_token_caveat` |
| 6.22 | فصل تحسن القائم عن الإضافة والتخارج | MT §5 | `growth` على `I`؛ `only_curr` و`only_prev` |
| 6.23 | «بدء إيراد» دون نسبة عند صفر الفترة السابقة | MT §5 | `revenue_start_ids` و`start_change` لكل سجل خارج النسبة؛ `V:pairwise_sets` |
| 6.24 | تغير التقييم منفصل وليس ربحًا نقديًا | MT §5 | `diff` على التقييم؛ التسمية؛ حكم المحلل |
| 6.25 | لا عائد كلي إلا بمدخلاته، والمعادلة المبسطة | MT §5 | `V:label_total_return` (مانع)؛ `V:narrative_claims` |
| 6.26 | منهج يراعي التدفقات وتوقيتها، وإلا ذكر التعذر | MT §5 | `derived_fields` للبسط والمقام؛ حكم المحلل |
| 6.27 | لا جمع لدخل سنة مع تغير تقييم فترة أخرى | MT §5 | `V:label_total_return`؛ حكم المحلل |
| 6.28 | لا تحويل للجزئي إلى سنوي محقق، والتقدير منفصل | MT §5 | `completeness: partial` يجعلها `not_comparable`؛ `V:growth_comparability_basis`؛ `assumptions` |
| 6.29 | الإشغال بمدخلات صحيحة فقط، والمساحي مستقل | MT §6 | مؤشرات مخصصة مشروطة؛ `data_requests` |
| 6.30 | التحصيل مع فصل المتأخرات القديمة والمقدم | MT §6 | حقول منفصلة؛ مؤشر مخصص |
| 6.31 | المتأخرات ونسبة المصروفات والإيراد للمتر بعد تماثل التعريفات | MT §6 | مؤشرات مخصصة؛ حكم المحلل |
| 6.32 | عدد الأصول المدرة ليس إشغالًا، والصفر لا يثبت شغورًا | MT §6 | تعريف `zero_rev_investment_value`؛ حكم المحلل |
| 6.33 | صافي الدخل التشغيلي ليس المتاح لخدمة الأيتام | MT §6 | تعريف `noi` القياسي؛ حكم المحلل |
| 6.34 | لا عدد أيتام إلا بتكلفة خدمة موثوقة ودخل متاح | MT §6 | حكم المحلل؛ `sources` و`assumptions` |

## 7. الدراسة العميقة وتشخيص الفجوات

| # | القاعدة | الملف | الإنفاذ |
|---|---|---|---|
| 7.1 | التحليل على مستوى المحفظة والمجموعات والأصول المؤثرة | DA §3 | `compute --filters --print`؛ `breakdowns` |
| 7.2 | تركّز القيمة والإيراد، والاعتماد على مصدر واحد | DA §4 | `value_concentration_top3` و`revenue_concentration_top3` (HHI) |
| 7.3 | الاستثماري غير المدر، والمنخفض قياسًا بهدف أو مقارنة | DA §4 | `zero_rev_investment_value` (`empty_as_zero`)؛ حكم المحلل |
| 7.4 | وزن الأصل في القيمة مقابل مساهمته في الإيراد | DA §4 | رسم مقارنة الأصول؛ `C:charts` |
| 7.5 | ضعف التأجير والتحصيل والمصروفات والعقود، عند وجود أدلة | DA §4 | مؤشرات مخصصة مشروطة؛ حكم المحلل |
| 7.6 | فرص الاستخدام والتأجير والترميم والتطوير المرحلي والشراكات | DA §4 | حكم المحلل؛ RS §7 |
| 7.7 | قدم التقييمات، ونقص البيانات، والتعارضات | DA §4 | `noi` غير قابل للاحتساب؛ `normalize:missing_classification` و`shared_deed` |
| 7.8 | عناصر سجل الفجوة | DA §5 | `schema` (`gap`)؛ `V:gap_evidence` (مانع)، `V:narrative_ids`، `V:top_lists`؛ `V:gap_completeness` |
| 7.9 | الاحتمال ليس حقيقة | DA §6 | `is_proven_problem` و`proven_cause`؛ `V:proven_problem_cause`، `V:gap_unsupported_claims`؛ `build` (الأحمر للمثبت فقط) |
| 7.10 | لا حكم بتدني الإيجار دون مقارنة مناسبة | DA §6 | حكم المحلل؛ `sources` |
| 7.11 | المراجع الخارجية من مصادر حديثة موثوقة، مع المصدر والتاريخ والحدود | DA §7 | `schema` (`web` يشترط `date` و`limits`)؛ حكم المحلل |
| 7.12 | لا متوسطات سوق مخترعة، ولا بيع أو تطوير من نسب عامة | DA §7 | `V:drastic_without_study` (مانع)؛ `V:drastic_recommendations`؛ حكم المحلل |

## 8. التقييم العام المختصر

| # | القاعدة | الملف | الإنفاذ |
|---|---|---|---|
| 8.1 | خلاصة من 120 إلى 180 كلمة بالعناصر الخمسة | DA §8 | `V:exec_summary_length` (مانع، بالطول بعد عرض الرموز)؛ `V:exec_summary_elements`؛ `V:hand_typed_numbers` |
| 8.2 | تقييم الأبعاد الأربعة منفصلة | DA §9 | `schema` (أربعة بنود بالضبط) |
| 8.3 | استخدام أهداف الجمعية المعتمدة إن توفرت | DA §9 | `criterion`؛ حكم المحلل |
| 8.4 | لا درجات ولا أحكام مطلقة بلا معيار | DA §9 | `V:narrative_scores` (مانع)؛ `V:assessment_scores`؛ `V:superlatives` |

## 9. التوصيات وخطة التنفيذ

| # | القاعدة | الملف | الإنفاذ |
|---|---|---|---|
| 9.1 | توصيات محددة مرتبطة بالفجوات | RS §2 | `V:recommendation_links` (مانع)؛ `V:gaps_without_recommendation` |
| 9.2 | الترتيب بالأثر والإلحاح والدليل والتنفيذ والتكلفة | RS §4 | `schema` (`ranking`)؛ `priority`؛ `V:recommendation_quality` |
| 9.3 | حقول التوصية الاثنا عشر | RS §3 | `V:recommendation_fields` (مانع) |
| 9.4 | خطة 30 و60 و90 يومًا | RS §5 | `V:plan_links` و`V:plan_completeness` (مانعان)؛ `C:structure` |
| 9.5 | مخرج مرحلي للمشروعات الطويلة، دون إيحاء بالانتهاء | RS §5 | `deliverable` (`schema`)؛ `V:plan_phasing` |
| 9.6 | لا عبارات عامة | RS §6 | `V:recommendation_quality`؛ حكم المحلل |
| 9.7 | لا بيع ولا هدم ولا تغيير استخدام لانخفاض العائد وحده | RS §6 | `V:drastic_without_study` (مانع)؛ `V:drastic_recommendations` |
| 9.8 | استمرارية خدمات المستفيدين | RS §6 | `V:drastic_recommendations` (أصل خدمي)؛ حكم المحلل |

## 10. فرص التحسين والسيناريوهات

| # | القاعدة | الملف | الإنفاذ |
|---|---|---|---|
| 10.1 | ثلاثة سيناريوهات عند توافر المدخلات | RS §7.1 | `schema` (`prepared` أو `not_prepared`) |
| 10.2 | افتراض بمصدر، وفصل التقديري عن المحقق | RS §7.2 | `V:scenarios` (مانع)؛ `build` (قسم السيناريوهات منفصل) |
| 10.3 | الإيراد والمصروف الإضافيان، وصافي الأثر، والرأسمالي، والتوقيت | RS §7.2 | `schema` (`scenario_case`)؛ `V:scenarios` (صافي الأثر)؛ `V:scenario_consistency` |
| 10.4 | أثر السنة الأولى مقابل الأثر السنوي | RS §7.2 | `year1_effect` و`run_rate_effect` |
| 10.5 | المتكرر منفصل عن التحصيل لمرة أو البيع | RS §7.2 | `recurring`؛ حكم المحلل |
| 10.6 | منع ازدواج الفرص على الأصل نفسه | RS §7.2 | `V:scenarios` (مانع)؛ `scenario_overlap_cross_type` |
| 10.7 | لا مبالغ اعتباطية، وتحديد المعلومات المطلوبة | RS §7.1 | `required_inputs` (`schema`)؛ `data_requests` |
| 10.8 | لا وعد مؤكد دون سند | RS §7.3 | `V:promised_outcomes`؛ حكم المحلل |

## 11. الداشبورد

| # | القاعدة | الملف | الإنفاذ |
|---|---|---|---|
| 11.1 | داشبورد فعلي عربي RTL متجاوب ومناسب للمجلس | DS §2 | `C:html_lang_dir`؛ `C:mobile_overflow` |
| 11.2 | HTML تفاعلي للتنزيل والفتح والطباعة، لا صورة ولا وصف | DS §2 | `build`؛ `C:offline`؛ `C:print` |
| 11.3 | تصميم هادئ: ألوان وخط وأرقام بفواصل | DS §3 | `build` (القالب)؛ مراجعة اللقطات |
| 11.4 | الأحمر للمثبت، والرمادي لغير المتاح، مع نص | DS §3 | `build`؛ `is_proven_problem` |
| 11.5 | التعريف | DS §4 | `build` (`meta`)؛ `C:structure` |
| 11.6 | التقييم التنفيذي المختصر | DS §4 | `C:structure` |
| 11.7 | ست إلى ثماني بطاقات بقائمتها | DS §4 | `V:card_count`؛ `C:cards` |
| 11.8 | أهم ثلاث فجوات وثلاثة إجراءات | DS §4 | `V:top_lists` و`V:plan_completeness` (مانعان)؛ `V:top_lists_count`؛ `C:structure` |
| 11.9 | الرسوم الأربعة، والزمني عند صلاحية المقارنة | DS §4 | `C:charts`؛ `trend.comparable` |
| 11.10 | جدول الأصول بالبحث والفرز والتصفية والتفاصيل والدليل | DS §4 | `C:table_search`، `C:table_sort`، `C:table_detail` |
| 11.11 | خطة 30 و60 و90 | DS §4 | `C:structure` |
| 11.12 | أقسام تفاصيل قابلة للفتح | DS §4 | `C:details_sections` |
| 11.13 | تحت بطاقة العائد: النطاق والأساس وتاريخ التقييم | DS §5 | `build`؛ `V:label_before_expenses` |
| 11.14 | «غير قابل للاحتساب بالبيانات الحالية» مع السبب، لا صفر | DS §5 | `V:metric_status_value` (مانع)؛ `C:cards`؛ `C:no_raw_tokens` |
| 11.15 | المرشحات: النوع والموقع والاستخدام والفترة | DS §6 | `filters`؛ `C:filters`؛ `C:period_filter` |
| 11.16 | مصدر حسابي واحد، ونطاق التصفية، وزر الإعادة | DS §6 | `C:parity`؛ `C:filters` |
| 11.17 | رسوم بسيطة بعناوين ووحدات، بلا ثلاثي الأبعاد | DS §7 | `build`؛ مراجعة اللقطات |
| 11.18 | لا خلط للعملات ولا لإيرادات غير متجانسة في مقارنة عوائد | DS §7 | `where` لمؤشر العائد؛ حكم المحلل |
| 11.19 | العرض التنفيذي للطباعة، والتفاصيل في أقسام | DS §8 | `C:print` |
| 11.20 | اختبار العربية والجوال والمرشحات والأزرار | DS §9 | `check_dashboard.mjs`؛ مراجعة اللقطات |
| 11.21 | لا زر تصدير أو تفاصيل لا يعمل | DS §9 | `C:buttons`؛ القالب بلا زر تصدير |
| 11.22 | لا بيانات اتصال أو معرفات شخصية | DS §10 | `normalize` (حذف `pii`)؛ `build` و`export` (حجب بالأنماط المشتركة، وفحص ما بعد الحفظ)؛ `V:pii_*` و`V:pii_export_patterns` (مانعة)؛ `V:pii_review_patterns`؛ `C:privacy` |

## 12. تحديث التحليل

| # | القاعدة | الملف | الإنفاذ |
|---|---|---|---|
| 12.1 | إعادة الفحص في كل إصدار، لا استبدال للأرقام | UV §1 | دورة كاملة؛ حكم المحلل |
| 12.2 | المقارنة بالمعرّفات، والفصل بين الفئات الأربع | UV §3 | `compare` |
| 12.3 | نسخة أحدث للسنة نفسها ليست نموًا | UV §4 | `compare` (`data_correction`، `growth_note`) |
| 12.4 | اختفاء صف لا يعني البيع | UV §4 | `compare` (`scope_change` «لا يُفترض بيعه») |
| 12.5 | لا ادعاء لذاكرة أو نسخة سابقة غير متاحة | UV §2 | `version_comparison: null`؛ حكم المحلل |
| 12.6 | تحديث متابعة التوصيات بالدليل أو الإفادة | UV §6 | `compare` («لم يبدأ» افتراضيًا، غير متحقق)؛ `recommendation_followup`؛ `export` (ورقة «متابعة التوصيات») |
| 12.7 | لا تنفيذ لمجرد تحسن رقم | UV §6 | `V:followup_status` |

## 13. المخرجات ومعيار الاكتمال

| # | القاعدة | الملف | الإنفاذ |
|---|---|---|---|
| 13.1 | رسالة مختصرة بالتقييم والأرقام والفجوات والتوصيات | DL §2 | `09_message.md`؛ قالب SK |
| 13.2 | الداشبورد التفاعلي الفعلي | DL §2 | `build`؛ `check_dashboard.mjs` |
| 13.3 | الدراسة التفصيلية والسجلات، في التبويبات أو مرفق | DL §2 | `export`؛ أقسام التفاصيل |
| 13.4 | خطة مرتبة بالأولوية، وقائمة البيانات الإضافية ومؤشراتها | DL §2 | `plan`؛ `data_requests` (`enables` في `schema`) |
| 13.5 | مطابقة البطاقات والرسوم والجداول للحسابات | DL §3 | `C:parity`، `C:cards`، `C:charts`؛ `V:metrics_reproducible` |
| 13.6 | تطابق البسط والمقام والفترة | DL §3 | `V:pairwise_sets`، `V:valuation_years_mixed`، `V:label_*` |
| 13.7 | عدم ازدواج الأصول والإيرادات | DL §3 | `V:duplicate_record_ids`، `V:total_rows_in_scope`، `V:isolated_in_metrics` |
| 13.8 | وضوح النواقص والافتراضات | DL §3 | `V:token_null_value`، `V:data_coverage_low`، `V:unknown_classification` |
| 13.9 | ربط التوصيات بالأدلة | DL §3 | `V:gap_evidence`، `V:recommendation_links`، `V:evidence_refs_other` |
| 13.10 | عمل التفاعل والطباعة | DL §3 | `C:buttons`، `C:table_*`، `C:print`، `C:mobile_overflow` |
| 13.11 | أدوات قراءة وحساب فعلية | DL §1 | حكم المحلل (السكربتات وPython) |
| 13.12 | ذكر العائق والبديل، ولا ادعاء | DL §5 | حكم المحلل (قوالب العائق) |
| 13.13 | البدء عند الاستلام وتنفيذ الدورة كاملة | DL §1 | تسلسل DL §4 |

## التحقق من المصفوفة

يجب أن يظهر كل وسم `[§x.y]` في الملف الأساسي المذكور. افحص ذلك بعد أي تعديل للمراجع:

```bash
cd "$SKILL_DIR/references" && python3 - <<'EOF'
import re, pathlib
files = {"SC": "scope-and-classification.md", "DC": "data-cleaning.md", "RF": "reference-file-controls.md", "MT": "metrics.md",
         "DA": "diagnosis-and-assessment.md", "RS": "recommendations-and-scenarios.md", "DS": "dashboard-spec.md",
         "UV": "updates-and-versioning.md", "DL": "delivery-checklist.md"}
rows = re.findall(r"^\| (\d+\.\d+) \| [^|]+ \| (\w\w)\b", pathlib.Path("traceability.md").read_text(encoding="utf-8"), re.M)
miss = [(i, f) for i, f in rows if f in files and not re.search(rf"§(?:[\d.]+\s*[·،–-]\s*)*{re.escape(i)}(?!\d)",
        pathlib.Path(files[f]).read_text(encoding="utf-8"))]
print(len(rows), "قاعدة؛ غير موسومة في ملفها:", miss or "لا شيء")
EOF
```
