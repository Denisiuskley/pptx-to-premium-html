    def generate_full_html(self, slides_data, stats) -> str:
        """Генерирует итоговый HTML-файл на основе встроенного шаблона и данных слайдов."""
        logger.info("Рендеринг HTML...")

        # Регулировка <title> и имени файла на основе данных докладчика
        speaker_name = slides_data[0].get("speaker_name", "") if slides_data else ""
        # Очистка имени от ВСЕХ маркеров форматирования для использования в заголовках и именах файлов
        clean_speaker = re.sub(r"\[\[.*?\]\]", "", speaker_name)
        
        # 1. Сначала определяем имя на основе докладчика
        if clean_speaker:
            safe_name = re.sub(r'[\\/*?:"<>|]', '', clean_speaker).strip()
            if safe_name:
                # Добавляем ПНИПУ только в название файла
                self.output_html = f"{safe_name} (ПНИПУ).html"
        
        # 2. Если имя все еще не определено (и не было передано снаружи), используем внутренний фолбэк
        if not self.output_html:
            self.output_html = "presentation_output.html"

        # --- Standalone / Embedding Mode ---
        is_standalone = DESIGN_CONFIG.get("STATIC_ASSETS_EMBED", False)
        head_part = BASE_HTML_TEMPLATE
        tail_part = BASE_HTML_TAIL
        
        # Регулировка <title> в head_part
        s_title = clean_speaker if clean_speaker else "Доклад"
        head_part = head_part.replace("{speaker_name}", s_title)

        # Logo path resolution
        logo_rel = DESIGN_CONFIG["paths"]["logo_white"]
        v_logo_path = BASE_DIR / "web_demo" / logo_rel
        if not v_logo_path.exists():
            v_logo_path = BASE_DIR / logo_rel
            
        if is_standalone:
            logger.info("Подготовка автономного (standalone) файла (все ресурсы вшиваются)...")
            # 1. Инлайним логотип
            logo = self._get_data_uri(v_logo_path)
            
            # 2. Инлайним шрифты
            fonts_css_path = BASE_DIR / "libs" / "fonts" / "fonts.css"
            fonts_inlined = inline_css_resources(fonts_css_path)
            head_part = re.sub(
                r'<link rel="stylesheet" href="libs/fonts/fonts.css">',
                f'<style>{fonts_inlined}</style>',
                head_part
            )
            
            # 3. Инлайним скрипты
            scripts_to_inline = [
                ("libs/gsap/gsap.min.js", r'<script src="libs/gsap/gsap.min.js" defer></script>'),
                ("libs/lucide/lucide.min.js", r'<script src="libs/lucide/lucide.min.js" defer></script>'),
                ("libs/mathjax/tex-mml-svg.js", r'<script src="libs/mathjax/tex-mml-svg.js" defer></script>')
            ]
            
            for rel_path, pattern in scripts_to_inline:
                s_path = BASE_DIR / rel_path
                s_content = get_file_content(s_path)
                # Robust replacement using lambda to avoid escape issues in large blocks
                head_part = re.sub(pattern, lambda m, c=s_content: f'<script>{c}</script>', head_part, flags=re.DOTALL)
        else:
            logo = DESIGN_CONFIG["paths"]["logo_white"]

        slides_content = ""
        total = len(slides_data)

        for idx, data in enumerate(slides_data):
            num = idx + 1
            title = data["title"]
            section_tag = self._generate_section_tag(data)
            text_panel_html = self._format_text_panel(data)

            # --- Умный шрифт: расчет panel_class на основе сложности контента ---
            total_text_chars = 0
            for item in data.get("content_items", []):
                if item["type"] == "text":
                    for p in item["data"]:
                        for s in p:
                            if isinstance(s, str):
                                # Считаем количество формул в сегменте
                                f_count = s.count("[[[MML_START]]]") + s.count("[[[MML_FB_START]]]")
                                if f_count > 0:
                                    # Очищаем сегмент от кода формул для честного подсчета текста
                                    # Ищем все вхождения [[[MML...]]] и удаляем их содержимое
                                    clean_s = re.sub(r"\[\[\[MML_START\]\].*?\[\[\[MML_END\]\]\]", "", s, flags=re.DOTALL)
                                    clean_s = re.sub(r"\[\[\[MML_FB_START\]\].*?\[\[\[MML_FB_END\]\]\]", "", clean_s, flags=re.DOTALL)
                                    # Также очищаем от новых маркеров форматирования
                                    clean_s = re.sub(r"\[\[[A-Z0-9_]+\]\]", "", clean_s)
                                    
                                    total_text_chars += len(clean_s)
                                    total_text_chars += (f_count * 20) # Significantly reduced weight for inline formulas
                                else:
                                    clean_s = re.sub(r"\[\[[A-Z0-9_]+\]\]", "", s)
                                    total_text_chars += len(clean_s)
                elif item["type"] == "table":
                    total_text_chars += 300  # Reduced table density impact
                elif item["type"] == "formula":
                    total_text_chars += 40 # Standalone formulas weight even less now


            panel_class = "analytical-panel animate-up"
            # -------------------------------------------------------------

            if is_standalone:
                embedded_count = 0
                for vis in data.get("visuals", []):
                    if vis.get("src") and not vis["src"].startswith("data:"):
                        # Картинки в web_demo/media/... относительны корня проекта в slides_data
                        v_path = BASE_DIR / "web_demo" / vis["src"]
                        if v_path.exists():
                            vis["src"] = self._get_data_uri(v_path)
                            embedded_count += 1
                        else:
                            logger.warning(f"Медиа не найдено для эмбеддинга: {v_path}")
                if embedded_count > 0:
                    logger.debug(f"Слайд {idx+1}: эмбедировано {embedded_count} изображений")
            
            if data["layout_type"] == "intro":
                slides_content += f"""
        <section class="slide hide-title">
            <div class="logo-container"><img src="{logo}" alt="Логотип" class="header-logo" onerror="this.style.display='none'; this.onerror=null;"></div>
            <div class="slide-header">
                <div class="slide-number">{num:02d} / {total:02d}</div>
            </div>
            <div class="slide-content-title">
                <h1 class="main-heading animate-up">{esc(title)}</h1>
            </div>
            <div class="presenter-card animate-up">
                <span class="presenter-label">Докладчик</span>
                <div class="presenter-name">{esc(data.get("speaker_name", ""))}</div>
                <div class="presenter-info">{esc(data.get("speaker_info", ""))}</div>
            </div>
        </section>"""

            elif data["layout_type"] == "conclusions_dual":
                slides_content += f"""
        <section class="slide">
            <div class="logo-container"><img src="{logo}" alt="Логотип" class="header-logo" onerror="this.style.display='none'; this.onerror=null;"></div>
            <div class="slide-header"><div class="slide-number">{num:02d} / {total:02d}</div><div class="slide-title">{esc(title)}</div></div>
            <div class="slide-split" style="grid-template-columns: 1fr 1fr; height: calc(100% - 145px);">
                <div class="{panel_class}">
                    {data["left_html"]}
                </div>
                <div class="{panel_class}">
                    {data["right_html"]}
                </div>
            </div>
        </section>"""

            elif data["layout_type"] == "research_roadmap":
                slides_content += f"""
        <section class="slide">
            <div class="logo-container"><img src="{logo}" alt="Логотип" class="header-logo" onerror="this.style.display='none'; this.onerror=null;"></div>
            <div class="slide-header"><div class="slide-number">{num:02d} / {total:02d}</div><div class="slide-title">{esc(title)}</div></div>
            <div class="slide-split" style="grid-template-columns: 1.4fr 0.8fr; height: calc(100% - 145px);">
                <div class="{panel_class}">
                    {data["content_html"]}
                </div>
                <div class="viz-card animate-up" style="display: flex; flex-direction: column; justify-content: center; align-items: center; background: radial-gradient(circle, var(--accent-soft) 0%, transparent 80%); border-radius: 3rem; padding: 3rem; border: 1px solid var(--glass-border); position: relative; overflow: hidden; height: 100%;">
                     <div class="rocket-glow"></div>
                     <i data-lucide="rocket" style="width: 120px; height: 120px; color: var(--accent); margin-bottom: 2rem; filter: drop-shadow(0 0 30px var(--accent)); transform: rotate(-45deg); animation: pulse-rocket 2s infinite ease-in-out;"></i>
                     <h4 style="font-family: 'Outfit'; font-size: var(--fs-research-year); margin: 0; font-weight: 800; background: linear-gradient(135deg, white 0%, var(--accent) 100%); -webkit-background-clip: text; -webkit-text-fill-color: transparent;">ROADMAP</h4>
                </div>
            </div>
        </section>"""

            elif data["layout_type"] == "full_text":
                slides_content += f"""
        <section class="slide">
            <div class="logo-container"><img src="{logo}" alt="Логотип" class="header-logo" onerror="this.style.display='none'; this.onerror=null;"></div>
            <div class="slide-header"><div class="slide-number">{num:02d} / {total:02d}</div><div class="slide-title">{esc(title)}</div></div>
            <div class="slide-content-full animate-up" style="width: 85%; margin: 2rem auto; height: calc(100% - 145px); min-height: 0;">
                <div class="{panel_class}" style="height: 100%;">
                    {text_panel_html}
                </div>
            </div>
        </section>"""

            else:
                # Стандартный макет со сплитом или только изображениями
                
                # Сортировка визуальных элементов
                data["visuals"] = self._spatial_sort(data["visuals"])
                
                # --- Расчет динамических пропорций (Dynamic Ratio) ---
                has_text = bool(text_panel_html.strip())
                base_ratio = DESIGN_CONFIG["layout"]["text_panel_ratio"] # 0.35
                layout_class = "slide-split"
                
                # 1. Предварительный расчет компоновки для определения количества колонок
                # Мы делаем "пробный" запуск с условной шириной, чтобы понять, пойдут ли картинки в один ряд или в несколько
                pre_results = self.get_best_layout(data["visuals"], 1200, 940)
                rows_pre, cols_pre = pre_results[0], pre_results[1]
                
                if has_text:
                    if cols_pre == 1:
                        # --- Адаптивный режим для ОДНОЙ КОЛОНКИ рисунков (1 шт или стек по вертикали) ---
                        grid_split = "1fr auto"
                        layout_class += " layout-auto-width"
                        cont_w = 1200 # Условное значение для расчета Area
                    else:
                        # --- Расчет по плотности для МНОГОКОЛОНОЧНЫХ компоновок ---
                        d_factor = max(0.65, min(1.2, total_text_chars / 1200.0 + 0.5))
                        dynamic_ratio = base_ratio * d_factor
                        dynamic_ratio = max(0.25, min(0.42, dynamic_ratio))
                        
                        grid_split = f"{dynamic_ratio:.3f}fr {1.0 - dynamic_ratio:.3f}fr"
                        cont_w = 1728 * (1.0 - dynamic_ratio)
                else:
                    cont_w = 1728
                    grid_split = "1fr"

                cont_h = 940 
                # Финальный расчет с уточненной шириной
                results = self.get_best_layout(data["visuals"], cont_w, cont_h)
                rows, cols, grid_styles, row_tmpl, col_tmpl = results

                # Сборка HTML визуальных элементов
                visuals_items_html = ""
                for idx, vis in enumerate(data["visuals"]):
                    style = grid_styles[idx] if idx < len(grid_styles) else ""
                    caption_html = (
                        f'<div class="viz-caption">{esc(vis.get("caption", ""))}</div>'
                        if vis.get("caption")
                        else ""
                    )
                    if vis.get("src"):
                        alt_text = (esc(vis["caption"]) if vis.get("caption") else "Изображение")
                        content_html = f'<div class="viz-box"><img src="{vis["src"]}" alt="{alt_text}"></div>'
                    else:
                        content_html = '<div class="error-box">Нет изображения</div>'
                    visuals_items_html += f'<div class="viz-item" style="{style}">{caption_html}{content_html}</div>'

                if has_text:
                    slides_content += f"""
        <section class="slide">
            <div class="logo-container"><img src="{logo}" alt="Логотип" class="header-logo" onerror="this.style.display='none'; this.onerror=null;"></div>
            <div class="slide-header"><div class="slide-number">{num:02d} / {total:02d}</div><div class="slide-title">{esc(title)}</div></div>
            <div class="slide-split {layout_class}" style="grid-template-columns: {grid_split}; height: calc(100% - 145px); min-height: 0;">
                <div class="{panel_class}">{text_panel_html}</div>
                <div class="img-stack animate-up" style="display: grid; grid-template-columns: {col_tmpl}; grid-template-rows: {row_tmpl}; gap: var(--gap-main);">{visuals_items_html}</div>
            </div>
        </section>"""
                else:
                    # Images only: full-width grid
                    slides_content += f"""
        <section class="slide">
            <div class="logo-container"><img src="{logo}" alt="Логотип" class="header-logo" onerror="this.style.display='none'; this.onerror=null;"></div>
            <div class="slide-header"><div class="slide-number">{num:02d} / {total:02d}</div><div class="slide-title">{esc(title)}</div></div>
            <div class="slide-content-full animate-up" style="width: 100%; margin: 0; height: calc(100% - 145px); display: grid; grid-template-columns: {col_tmpl}; grid-template-rows: {row_tmpl}; gap: var(--gap-main); padding: 0;">
                {visuals_items_html}
            </div>
        </section>"""

        logger.info(f"Сохранение в {self.output_html}...")
        with open(self.output_html, "w", encoding="utf-8") as f:
            f.write(head_part)
            f.write(slides_content)
            f.write(tail_part)

        logger.info("\n" + "=" * 40)
        logger.info(" ИТОГОВЫЙ ОТЧЕТ КОНВЕРТАЦИИ")
        logger.info("=" * 40)
        logger.info(f" Всего слайдов:   {stats['total_slides']}")
        logger.info(f" Успешных фото:   {stats['images_ok']}")
        logger.info(f" Таблиц:          {stats['tables']}")
        logger.info(f" Формул:          {stats['formulas']}")
        logger.info(f" Пропущено:       {stats['images_fail']} (см. лог выше)")
        logger.info("=" * 40)

    def _format_text_panel(self, slide_info: dict) -> str:
        """Формирует HTML-панель с чередованием текста, таблиц и формул."""
        parts = []
        bullet_config = DESIGN_CONFIG.get("bullet_lists", {})
        bullet_icon_map = bullet_config.get("icon_map", {})
        bullet_icon_size = DESIGN_CONFIG["icon_size"]
        bullet_indent = bullet_config.get("indent", "2rem")
        bullet_border = bullet_config.get("border_left", "2px solid rgba(0,242,255,0.15)")
        bullet_bg = bullet_config.get("background", "rgba(0,242,255,0.03)")

        # Итерируемся по объединенному списку контента для соблюдения порядка
        for item_block in slide_info.get("content_items", []):
            item_type = item_block["type"]
            data = item_block["data"]

            if item_type == "text":
                # Обработка сегментированных параграфов
                for p in data:
                    # 1. Сбор сегментов и замена формул на токены
                    segments_to_join = []
                    formula_store = {}
                    
                    if isinstance(p, list):
                        for idx, seg in enumerate(p):
                            if seg.startswith("[[[MML_START]]]"):
                                token = f"[[_F{idx}_]]"
                                mathml = seg[len("[[[MML_START]]]"): -len("[[[MML_END]]]")]
                                formula_store[token] = f'<span class="formula-container">{mathml}</span>'
                                segments_to_join.append(token)
                            elif seg.startswith("[[[MML_FB_START]]]"):
                                token = f"[[_FB{idx}_]]"
                                fb = seg[len("[[[MML_FB_START]]]"): -len("[[[MML_FB_END]]]")]
                                formula_store[token] = f'<span class="formula-fallback">{esc(fb)}</span>'
                                segments_to_join.append(token)
                            else:
                                segments_to_join.append(seg)
                    else:
                        segments_to_join.append(str(p))

                    full_content_with_tokens = "".join(segments_to_join)
                    if not full_content_with_tokens.strip():
                        continue

                    # 2. Очистка для поиска буллитов (наши токены выживут)
                    clean_search_text = re.sub(r"<[^>]+>", "", full_content_with_tokens)
                    items = self._split_text_into_items(clean_search_text)

                    for item in items:
                        # 3. Логика распознавания «Вывода» и подготовка текста
                        raw_text = item["text"].strip()
                        is_conclusion = raw_text.lower().startswith("вывод:")
                        
                        if is_conclusion:
                            # Очищаем от префикса «Вывод:»
                            item_text = re.sub(r"^[Вв]ывод[:\s]+", "", raw_text).strip()
                        else:
                            item_text = item["text"]

                        display_text = html.escape(item_text)
                        
                        # Возврат форматирования из маркеров
                        display_text = display_text.replace("[[SUB_S]]", "<sub>").replace("[[SUB_E]]", "</sub>")
                        display_text = display_text.replace("[[SUP_S]]", "<sup>").replace("[[SUP_E]]", "</sup>")
                        display_text = display_text.replace("[[B_S]]", "<strong>").replace("[[B_E]]", "</strong>")
                        display_text = display_text.replace("[[I_S]]", "<em>").replace("[[I_E]]", "</em>")

                        for token, real_html in formula_store.items():
                            display_text = display_text.replace(token, real_html)
                        
                        if is_conclusion:
                            display_text = f"<strong>{display_text}</strong>"
                        
                        if item.get("is_bullet") or is_conclusion:
                            if is_conclusion:
                                icon = "check-check"
                            else:
                                keyword_icon = self.get_icon_by_text(item["text"])
                                default_bullet = "diamond"
                                icon = keyword_icon if keyword_icon != "chevron-right" else default_bullet
                            
                            # Применяем вытягивание ромба через спец. класс
                            marker_class = "animate-marker"
                            if icon == "diamond":
                                marker_class += " marker-stretched"

                            # Специальная верстка для блока "Вывод:" с акцентной полосой
                            wrapper_class = "list-item-bullet animate-up"
                            wrapper_style = f"padding-left: {bullet_indent};"
                            if is_conclusion:
                                wrapper_class += " list-item-conclusion"
                            else:
                                wrapper_style += f" border-left: {bullet_border}; background: {bullet_bg};"

                            parts.append(
                                f'<div class="{wrapper_class}" style="{wrapper_style}">'
                                f'<i data-lucide="{icon}" class="{marker_class}" style="width: {bullet_icon_size}; height: {bullet_icon_size}; flex-shrink: 0;"></i>'
                                f'<div class="list-text">{display_text}</div></div>'
                            )
                        else:
                            keyword_icon = self.get_icon_by_text(item["text"])
                            icon = keyword_icon if keyword_icon != "chevron-right" else "diamond"
                            
                            marker_class = "animate-marker"
                            if icon == "diamond":
                                marker_class += " marker-stretched"

                            parts.append(
                                f'<div class="list-item animate-up">'
                                f'<i data-lucide="{icon}" class="{marker_class}" style="width: {DESIGN_CONFIG["icon_size"]}; height: {DESIGN_CONFIG["icon_size"]}; flex-shrink: 0;"></i>'
                                f'<div class="list-text">{display_text}</div></div>'
                            )

            elif item_type == "table":
                parts.append(str(data))

            elif item_type == "formula":
                parts.append(f'<div class="formula-block animate-up">{data}</div>')

        return "".join(parts)

    def _split_text_into_items(self, text: str) -> list:
        """Разбирает многострочный текст на отдельные пункты.
        Возвращает список dict: [{"text": str, "is_bullet": bool, "bullet_char": str|None, "level": int}]
        Префиксы списков (цифры, bullets, дефисы) удаляются из текста, но тип маркера запоминается.
        """
        if not text:
            return []
        lines = text.split("\n")
        items = []
        # Символы маркера из конфига (ключи словаря)
        bullet_chars = set(
            DESIGN_CONFIG.get("bullet_lists", {}).get("icon_map", {}).keys()
        )
        # Расширенный паттерн для удаления префиксов: цифры с точками, bullets, дефисы
        # Исключаем буквы из длинных слов ([a-zA-Zа-яА-Я]+), оставляя только одиночные буквы-маркеры (a., B., а.)
        prefix_pattern = r"^\s*(\d+[\.\)]\s*|[a-zа-я][\.\)]\s*|[•◦▪\-–—*·ü]\s*)+\s*"
        for line in lines:
            stripped = line.strip()
            if not stripped:
                continue
            first_char = stripped[0]
            # КОРРЕКЦИЯ: Более гибкое определение списков (любая нумерация или пунктировка)
            is_bullet = (
                (first_char in bullet_chars) or 
                bool(re.match(r"^\s*(\d+[\.\)]|[a-zA-Zа-яА-Я][\.\)]|\([\d\w]\))\s*", stripped)) or
                bool(re.match(r"^\s*[IVXLCDMivxlcdm]+\.\s*", stripped))
            )
            
            cleaned = re.sub(prefix_pattern, "", stripped).strip()
            if cleaned:
                items.append(
                    {
                        "text": cleaned,
                        "is_bullet": is_bullet,
                        "bullet_char": first_char if is_bullet else None,
                        "level": 0,
                    }
                )
        return items

    def get_icon_by_text(self, text: str) -> str:
        """Возвращает идентификатор иконки на основе текста (по ключевым словам)."""
        text = text.lower()
        for icon, keywords in DESIGN_CONFIG["icon_mapping"].items():
            if any(kw in text for kw in keywords):
                return icon
        return "chevron-right"

    def _generate_section_tag(self, data: dict) -> str:
        """Генерирует тег секции (например, "Данные", "Результаты") на основе текста."""
        all_texts = []
        for block in data.get("content_items", []):
            if block["type"] == "text":
                for p in block["data"]:
                    all_texts.append(" ".join(s for s in p if not s.startswith("[[[") ))
        
        text_combined = " ".join(all_texts).lower()
        if any(kw in text_combined for kw in ["данные", "таблиц", "исходн"]):
            return "Данные"
        if any(kw in text_combined for kw in ["результат", "вывод", "итог"]):
            return "Результаты"
        if any(kw in text_combined for kw in ["описание", "метод", "подход"]):
            return "Описание"
        if any(kw in text_combined for kw in ["анализ", "исследовани"]):
            return "Анализ"
        return "Описание"

    def call_ai(self, prompt: str) -> str | None:
        """Вызывает LLM через OpenRouter API с кэшированием."""
        env = self.load_env()
        api_key = env.get("OPENROUTER_API_KEY")
        model = env.get("MODEL", DESIGN_CONFIG["ai_config"]["default_model"])

        if not api_key:
            logger.warning("OPENROUTER_API_KEY не найден в .env")
            return None

        cache = LLMCache()
        cached = cache.get(prompt, model)
        if cached:
            logger.info(f"[LLM] Ответ получен из кэша для prompt: {prompt[:50]}...")
            return cached

        headers = {
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
        }
        payload = {"model": model, "messages": [{"role": "user", "content": prompt}]}

        try:
            response = requests.post(
                "https://openrouter.ai/api/v1/chat/completions",
                headers=headers,
                data=json.dumps(payload),
                timeout=30,
            )
            if response.status_code == 200:
                result = response.json()
                content = result["choices"][0]["message"]["content"]
                cache.set(prompt, model, content)
                return content
            else:
                logger.error(
                    f"Ошибка API: {response.status_code} - {response.text[:200]}"
                )
        except Exception as e:
            logger.error(f"Исключение при вызове ИИ: {e}")
        return None

    def process_txt_files(self) -> None:
        """Обрабатывает вспомогательные текстовые файлы с использованием AI для структурирования."""
        files = {
            "Выводы.txt": {
                "type": "conclusions",
                "prompt": "Проанализируй текст и извлеки ключевые факты. СТРОГАЯ СТРУКТУРА: 1. Выполненные работы (МАКСИМУМ 6 пунктов): перечисли основные этапы и действия. 2. Результаты (МАКСИМУМ 6 пунктов): перечисли конкретные выводы, достижения и показатели. ПРАВИЛА: - СТРОГО соблюдай хронологический порядок событий. - Текст должен быть максимально сжатым, строгим и технически точным. - Объединяй весь опыт, не пропуская значимых деталей (даты, цифры), но формулируй их крайне емко. - Удали вводные слова, пояснения и 'воду'. ФОРМАТ: Верни только JSON-массив из кратких тезисов (сначала до 6 пунктов по работам, затем до 6 пунктов по результатам). Текст: {content}",
            },
            "Направление дальнейших исследований.txt": {
                "type": "research",
                "prompt": "Извлеки основные направления дальнейших исследований. ПОРЯДОК: Сначала планируемые действия, затем ожидаемые эффекты и цели. ПРАВИЛА: - Соблюдай логическую и хронологическую последовательность. - Максимально краткий и сухой научный стиль. - Сохрани всю фактологическую базу данных. ФОРМАТ: Верни только JSON-массив строк. Текст: {content}",
            },
        }

        for filename, config in files.items():
            if not os.path.exists(filename):
                continue
            logger.info(f"Обработка текстового файла: {filename}")
            with open(filename, "r", encoding="utf-8") as f:
                content = f.read()

            items = []
            ai_used = False
            api_key = self.load_env().get("OPENROUTER_API_KEY")

            if api_key:
                prompt = config["prompt"].format(content=content[:8000])
                ai_response = self.call_ai(prompt)
                if ai_response:
                    try:
                        # Clean markdown code fences if present
                        cleaned = ai_response.strip()
                        if cleaned.startswith("```json"):
                            cleaned = cleaned.split("```json", 1)[1]
                        if cleaned.startswith("```"):
                            cleaned = cleaned.split("```", 1)[1]
                        if "```" in cleaned:
                            cleaned = cleaned.split("```", 1)[0]
                        cleaned = cleaned.strip()

                        parsed = json.loads(cleaned)
                        if isinstance(parsed, list) and all(
                            isinstance(x, str) for x in parsed
                        ):
                            items = [p.strip() for p in parsed if p.strip()]
                            ai_used = True
                            logger.info(
                                f"[AI] Получено {len(items)} пунктов из {filename}"
                            )
                        else:
                            logger.warning(
                                f"[AI] Ответ не является массивом строк, используем fallback"
                            )
                    except json.JSONDecodeError as e:
                        logger.warning(
                            f"[AI] Не удалось распарсить JSON: {e}, используем fallback"
                        )

            if not items:
                # Fallback: разбиваем по строкам
                paragraphs = [p.strip() for p in content.split("\n") if p.strip()]
                items = paragraphs
                if not ai_used:
                    logger.info(
                        f"[Fallback] Использовано разбиение по строкам для {filename} ({len(items)} пунктов)"
                    )

            if config["type"] == "conclusions":
                processed_items = []
                for item in items:
                    clean_item = clean_text(item)
                    icon = self.get_icon_by_text(clean_item)
                    processed_items.append(
                        f"<div class='summary-item animate-up'><i data-lucide='{icon}' class='animate-marker' style='width: 1.6rem; height: 1.6rem; flex-shrink: 0;'></i> <div class='list-text'>{esc(clean_item)}</div></div>"
                    )

                mid = (len(processed_items) + 1) // 2
                left_html = "".join(processed_items[:mid])
                right_html = "".join(processed_items[mid:])

                slides_data.append(
                    {
                        "title": "Выводы и результаты этапа",
                        "layout_type": "conclusions_dual",
                        "left_html": left_html,
                        "right_html": right_html,
                        "content_items": [],
                    }
                )
            elif config["type"] == "research":
                items_html = []
                for item in items:
                    clean_item = clean_text(item)
                    items_html.append(
                        f"<div class='roadmap-item animate-up'><i data-lucide='rocket' class='animate-marker' style='width: 1.6rem; height: 1.6rem; flex-shrink: 0;'></i> <div class='list-text'>{esc(clean_item)}</div></div>"
                    )
                slides_data.append(
                    {
                        "title": "Направление дальнейших исследований",
                        "layout_type": "research_roadmap",
                        "content_html": "".join(items_html),
                        "content_items": [],
                    }
                )

    def load_env(self) -> dict:
        """Загружает переменные окружения из .env файла."""
        env = {}
        path = DESIGN_CONFIG["ai_config"]["env_path"]
        if os.path.exists(path):
            with open(path, "r", encoding="utf-8") as f:
                for line in f:
                    line = line.strip()
                    if "=" in line and not line.startswith("#"):
                        k, v = line.split("=", 1)
                        env[k.strip()] = v.strip().strip('"').strip("'")
        return env
