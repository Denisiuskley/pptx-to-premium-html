# ==============================================================================
# ВСТРОЕННЫЙ HTML-ШАБЛОН (все стили и разметка вынесены из основного скрипта)
# ==============================================================================
BASE_HTML_TEMPLATE = """<!DOCTYPE html>
<html class="no-js" lang="ru">

<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Доклад: {speaker_name} (ПНИПУ)</title>
    <link rel="stylesheet" href="libs/fonts/fonts.css">
    <script src="libs/gsap/gsap.min.js" defer></script>
    <script src="libs/lucide/lucide.min.js" defer></script>
    <!-- Local MathJax configuration -->
    <script>
        window.MathJax = {
            options: {
                enableExplorer: false, // Отключаем Explorer, который может перекрывать формулы желтыми плашками
                enableAssistiveMml: true
            },
            loader: {
                load: ['[tex]/ams']
            },
            svg: {
                fontCache: 'global',
                scale: 1.2
            },
            startup: {
                ready: () => {
                    console.log('MathJax is ready (SVG mode)!');
                    MathJax.startup.defaultReady();
                    MathJax.startup.promise.then(() => {
                        // Анимация появления формул после рендеринга
                        gsap.to('.formula-container, .formula-block', {
                            opacity: 1,
                            y: 0,
                            duration: 0.8,
                            stagger: 0.1,
                            ease: 'power2.out'
                        });
                    });
                }
            }
        };
    </script>
    <script src="libs/mathjax/tex-mml-svg.js" defer></script>
        <style>
            :root {
                --bg-deep: #05080e;
                --bg-card: rgba(15, 23, 42, 0.7);
                --accent: #00f2ff;
                --accent-soft: rgba(0, 242, 255, 0.1);
                --text-main: #f8fafc;
                --text-dim: #cbd5e1;
                --glass-border: rgba(255, 255, 255, 0.1);
                --gradient-accent: linear-gradient(135deg, #00f2ff 0%, #0066ff 100%);
                --glow-cyan: 0 0 20px rgba(0, 242, 255, 0.4);

                --logo-height: 160px;
                --logo-top: -1.0rem;
                --logo-right: 4rem;
                --header-height: 75px;
                --slide-padding-v: 2rem;
                --slide-padding-h: 4rem;
                --panel-padding: 2.0rem;

                --fs-main-title: 4.2rem;
                --fs-slide-title: 1.4rem;
                --fs-slide-num: 1.1rem;
                --fs-sub-heading: 1.8rem;
                --fs-text-main: 1.3rem;
                --fs-tag: 0.9rem;
                --fs-viz-caption: 1.1rem;
                --fs-viz-desc: 1.0rem;
                --fs-presenter-name: 1.5rem;
                --fs-presenter-info: 1.1rem;
                --fs-presenter-label: 0.8rem;
                --fs-research-year: 5rem;

                /* Таблицы */
                --fs-table-th: 1.05rem;
                --fs-table-td: 1.1rem;
                --fs-formula: 1.25rem;
                --gap-main: 2rem;
                --gap-items: 1.2rem;

                /* Adaptive Squeeze Factors */
                --squeeze-factor: 1;
                --base-gap: 0.5rem;
                --base-margin: 0.5rem;
                --base-lh: 1.6;

                /* Config-driven variables */
                --icon-size: 1.6rem;
                --bullet-icon-size: 1.6rem;
                --bullet-indent: 2rem;
                --bullet-border: 2px solid rgba(0, 242, 255, 0.15);
                --bullet-bg: rgba(0, 242, 255, 0.03);
                --formula-fallback-font: 'Roboto Mono', monospace;
                --formula-fallback-size: 1.1em;
            }

        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
            -webkit-font-smoothing: antialiased;
        }

        body {
            background-color: var(--bg-deep);
            color: var(--text-main);
            font-family: 'Inter', sans-serif;
            overflow: hidden;
            height: 100vh;
        }

        .presentation {
            position: relative;
            width: 100vw;
            height: 100vh;
            display: flex;
            transition: transform 0.8s cubic-bezier(0.85, 0, 0.15, 1);
        }

        .slide {
            flex: 0 0 100vw;
            height: 100vh;
            display: flex;
            flex-direction: column;
            padding: var(--slide-padding-v) var(--slide-padding-h);
            position: relative;
            overflow: hidden;
        }

        .bg-element {
            position: absolute;
            z-index: -1;
            filter: blur(80px);
            opacity: 0.15;
            border-radius: 50%;
        }

        .bg-1 {
            width: 600px;
            height: 600px;
            background: var(--gradient-accent);
            top: -200px;
            right: -100px;
        }

        .bg-2 {
            width: 400px;
            height: 400px;
            background: #ff00ea;
            bottom: -100px;
            left: -100px;
            opacity: 0.05;
        }

        .slide-header {
            display: flex;
            justify-content: flex-start;
            align-items: center;
            gap: 1.5rem;
            margin-bottom: var(--margin-header);
            border-bottom: 1px solid var(--glass-border);
            padding-bottom: 0.5rem;
            position: relative;
            height: var(--header-height);
        }

        .logo-container {
            position: absolute;
            top: var(--logo-top);
            right: var(--logo-right);
            height: var(--logo-height);
            display: flex;
            align-items: center;
            z-index: 10;
        }

        .header-logo {
            height: 100%;
            width: auto;
            filter: drop-shadow(0 0 20px rgba(0, 242, 255, 0.5));
            opacity: 0.9;
        }

        .hide-title .slide-title {
            display: none;
        }

        .slide-title {
            font-family: 'Outfit', sans-serif;
            font-size: var(--fs-slide-title);
            letter-spacing: 0.05em;
            text-transform: uppercase;
            color: var(--accent);
            font-weight: 700;
        }

        .slide-number {
            font-family: 'Roboto Mono', monospace;
            font-size: var(--fs-slide-num);
            color: var(--text-dim);
            padding-right: 1rem;
            border-right: 1px solid var(--glass-border);
        }

        .slide-content-title {
            flex: 1;
            display: flex;
            flex-direction: column;
            justify-content: center;
            max-width: 1400px;
            margin: 0 auto;
            width: 100%;
        }

        .main-heading {
            font-family: 'Inter', sans-serif;
            font-size: var(--fs-main-title);
            font-weight: 800;
            line-height: 1.1;
            margin-bottom: 3rem;
            text-align: center;
            background: linear-gradient(to bottom right, #fff 50%, #94a3b8);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
        }

        .presenter-card {
            background: var(--bg-card);
            backdrop-filter: blur(12px);
            border: 1px solid var(--glass-border);
            padding: 2rem;
            border-right: 4px solid var(--accent);
            align-self: flex-start;
            width: 500px;
            position: relative;
            z-index: 20;
        }

        .presenter-label {
            color: var(--accent);
            font-size: var(--fs-presenter-label);
            text-transform: uppercase;
            letter-spacing: 0.2em;
            margin-bottom: 0.5rem;
            display: block;
        }

        .presenter-name {
            font-size: var(--fs-presenter-name);
            font-weight: 800;
            margin-bottom: 0.25rem;
            color: var(--text-main);
        }

        .presenter-info {
            color: var(--text-dim);
            font-size: var(--fs-presenter-info);
            line-height: 1.6;
            word-wrap: break-word;
            overflow-wrap: anywhere;
            max-width: 90%;
        }

        .slide.hide-title .presenter-card {
            position: absolute;
            bottom: var(--slide-padding-v);
            left: var(--slide-padding-h);
        }

        .slide-split {
            display: grid;
            grid-template-columns: 1.2fr 1.8fr;
            gap: var(--gap-main);
            flex: 1;
            min-height: 0;
        }

        .img-stack {
            display: flex;
            flex-direction: row;
            gap: var(--gap-main);
            align-items: stretch;
            height: 100%;
            min-height: 0;
            width: 100%;
            overflow: hidden; /* Firefox fix for image clipping */
        }

        .formula-block {
            margin: 1.5rem 0;
            text-align: center;
            background: rgba(0, 242, 255, 0.02);
            padding: 1rem;
            border-radius: 8px;
            border: 1px solid var(--glass-border);
        }

        .slide-split.layout-auto-width {
            grid-template-columns: minmax(25%, 1fr) auto !important;
        }
        .layout-auto-width .img-stack {
            width: auto;
            justify-content: flex-end;
        }
        .layout-auto-width .viz-item {
            width: auto;
            flex: none;
        }
        .layout-auto-width .viz-box {
            width: auto;
            max-width: 65vw;
        }
        .layout-auto-width .viz-box img {
            width: 100%;
            height: 100%;
            max-width: 100%;
            max-height: 100%;
            min-width: 0;
            min-height: 0;
            object-fit: contain;
        }

        .viz-item {
            display: flex;
            flex-direction: column;
            gap: 1rem;
            flex: 1;
            min-height: 0;
        }

        .viz-caption {
            border-left: 4px solid var(--accent);
            padding: calc(0.5rem * var(--squeeze-factor)) calc(1rem * var(--squeeze-factor));
            background: linear-gradient(to right, var(--accent-soft), transparent);
            color: var(--text-main);
            font-size: var(--fs-viz-caption);
            font-weight: 600;
            line-height: 1.3;
        }

        .data-table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 1rem;
            font-size: var(--fs-table-td);
            color: var(--text-main);
            font-family: 'Inter', sans-serif;
        }
        .data-table th, .data-table td {
            border: 1px solid var(--glass-border);
            padding: 0.25rem 0.5rem;
            text-align: left;
            vertical-align: top;
            line-height: 1.3;
            word-break: break-word;
            overflow-wrap: break-word;
        }
        .data-table th {
            background: var(--accent-soft);
            color: var(--accent);
            font-weight: 700;
            letter-spacing: 0.05em;
            font-size: var(--fs-table-th);
        }
        .data-table tr:nth-child(even) {
            background: rgba(255, 255, 255, 0.02);
        }
        .data-table tr:hover {
            background: rgba(0, 242, 255, 0.05);
        }

        .analytical-panel {
            background: var(--bg-card);
            backdrop-filter: blur(8px);
            border: 1px solid var(--glass-border);
            padding: var(--panel-padding);
            border-radius: 8px;
            display: flex;
            flex-direction: column;
            gap: calc(var(--base-gap) * var(--squeeze-factor));
            overflow-y: hidden;
            min-height: 0; /* Firefox fix */
            box-shadow: inset 0 1px 0 rgba(255,255,255,0.1), inset 1px 0 0 rgba(255,255,255,0.05), 0 8px 32px rgba(0,0,0,0.3);
            position: relative;
        }

        .section-tag {
            font-family: 'Roboto Mono', monospace;
            background: var(--accent-soft);
            color: var(--accent);
            padding: 4px 12px;
            font-size: var(--fs-tag);
            width: fit-content;
            border-radius: 4px;
        }

        .sub-heading {
            font-size: var(--fs-sub-heading);
            font-weight: 600;
            margin-bottom: calc(var(--base-margin) * var(--squeeze-factor));
        }

        .list-item {
            display: flex;
            gap: 1rem;
            margin-bottom: calc(var(--base-margin) * var(--squeeze-factor));
            align-items: flex-start;
        }

        .list-item i, .list-item svg {
            margin-top: 4px;
            color: var(--accent);
            flex-shrink: 0;
            width: var(--icon-size);
            height: var(--icon-size);
            transition: transform 0.3s cubic-bezier(0.34, 1.56, 0.64, 1);
        }

        .animate-marker {
            opacity: 0;
            transform: scale(0);
            display: inline-block;
        }

        .marker-pulse {
            animation: markerBreathe 3s infinite ease-in-out;
        }

        @keyframes markerBreathe {
            0%, 100% { transform: scale(1); filter: drop-shadow(0 0 0px var(--accent)); opacity: 0.8; }
            50% { transform: scale(1.15); filter: drop-shadow(0 0 8px var(--accent)); opacity: 1; }
        }

        .list-item-bullet {
            display: flex;
            gap: 1rem;
            margin-bottom: calc(var(--base-margin) * var(--squeeze-factor));
            align-items: flex-start;
            padding-left: var(--bullet-indent);
            border-left: var(--bullet-border);
            background: var(--bullet-bg);
            border-radius: 0 8px 8px 0;
            padding: 0.5rem 1rem 0.5rem 0.5rem;
        }
        .list-item-bullet i, .list-item-bullet svg {
            margin-top: 4px;
            color: var(--accent);
            flex-shrink: 0;
            width: var(--bullet-icon-size);
            height: var(--bullet-icon-size);
            opacity: 0.85;
        }

        .list-item-conclusion {
            border-left: 5px solid var(--accent) !important;
            background: linear-gradient(90deg, var(--accent-soft), transparent) !important;
            padding: 1rem 1rem 1rem 1.2rem !important;
            border-radius: 6px 16px 16px 6px !important;
            margin-top: 1.2rem;
            margin-bottom: 2rem;
            position: relative;
            overflow: visible;
        }

        .list-item-conclusion i, .list-item-conclusion svg {
            animation: pulse-rocket 2s infinite ease-in-out;
            filter: drop-shadow(0 0 8px var(--accent));
        }

        .list-item-conclusion::after {
            content: '';
            position: absolute;
            left: -5px;
            top: 20%;
            height: 60%;
            width: 8px;
            background: var(--accent);
            filter: blur(8px);
            opacity: 0.6;
            pointer-events: none;
        }

        .viz-container:has(.viz-item) {
            max-width: fit-content;
        }
         .viz-container:has(.viz-item) .viz-item {
             width: 100%;
         }

         .viz-item.wide {
             grid-column: span 2;
         }
         .viz-item.tall {
             grid-row: span 2;
         }

        .list-text {
            color: var(--text-dim);
            font-size: var(--fs-text-main);
            line-height: calc(var(--base-lh) * var(--squeeze-factor));
        }

        .list-text strong {
            color: var(--text-main);
        }

        .formula-fallback {
            font-family: var(--formula-fallback-font, 'Roboto Mono'), monospace;
            font-size: var(--formula-fallback-size, 1.1em);
            white-space: nowrap;
            color: var(--accent);
            opacity: 0.9;
            display: inline-block;
            padding: 0.1rem 0.3rem;
            background: rgba(0, 242, 255, 0.05);
            border-radius: 4px;
        }

        .formula-container {
            display: inline-block;
            vertical-align: middle;
            margin: 0 0.3rem;
            opacity: 0;
            transform: translateY(5px);
        }

        .formula-block {
            background: var(--bg-card);
            backdrop-filter: blur(12px);
            border: 1px solid var(--glass-border);
            padding: 1rem 1.5rem;
            margin: 0.75rem 0;
            display: flex;
            justify-content: center;
            align-items: center;
            box-shadow: 0 12px 40px rgba(0,0,0,0.4);
            opacity: 0;
            transform: translateY(10px);
            transition: border-color 0.4s, box-shadow 0.4s;
        }

        .formula-block:hover {
            border-color: var(--accent);
            box-shadow: 0 0 30px rgba(0, 242, 255, 0.15);
        }

        mjx-container[jax="SVG"] {
            color: var(--text-main);
            margin: 0 !important;
            padding: 0 !important;
        }

        .formula-container mjx-container[jax="SVG"] {
            display: inline-block !important;
            vertical-align: middle !important;
            base-line: middle !important;
        }

        .marker-stretched {
            transform: scaleX(1.5);
            display: inline-block;
        }

        .viz-box {
            background: radial-gradient(circle at center, rgba(0, 242, 255, 0), transparent);
            border: 1px solid rgba(255, 255, 255, 0.1);
            border-radius: 12px;
            display: flex;
            align-items: center;
            justify-content: center;
            position: relative;
            overflow: hidden;
            height: 100%;
            width: 100%;
            min-height: 0;
            cursor: zoom-in;
            transition: all 0.5s cubic-bezier(0.4, 0, 0.2, 1);
        }

        .viz-box:hover {
            border-color: var(--accent);
            box-shadow:
                0 0 20px rgba(0, 242, 255, 0.2),
                inset 0 0 30px rgba(0, 242, 255, 0.15);
        }

        .viz-box::after {
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            background-image:
                linear-gradient(rgba(0, 242, 255, 0.05) 1px, transparent 1px),
                linear-gradient(90deg, rgba(0, 242, 255, 0.05) 1px, transparent 1px);
            background-size: 20px 20px;
            opacity: 0;
            transition: opacity 0.4s;
            pointer-events: none;
        }

        .viz-box:hover::after {
            opacity: 1;
        }

        .viz-box img {
            width: 100%;
            height: 100%;
            max-width: 100%;
            max-height: 100%;
            min-width: 0;
            min-height: 0;
            object-fit: contain;
            display: block;
            transition: transform 0.6s cubic-bezier(0.4, 0, 0.2, 1);
        }

        .viz-box:hover img {
            transform: scale(1.02);
        }

        .lightbox {
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background: rgba(5, 8, 14, 0.95);
            backdrop-filter: blur(15px);
            z-index: 2000;
            display: none;
            justify-content: center;
            align-items: center;
            opacity: 0;
        }

        .lightbox-content {
            width: 90vw;
            height: 85vh;
            position: relative;
            display: flex;
            justify-content: center;
            align-items: center;
            box-shadow: 0 0 50px rgba(0, 0, 0, 0.5);
            border: 1px solid var(--glass-border);
            background: rgba(15, 23, 42, 0.8);
            border-radius: 12px;
            padding: 1rem;
        }

        .lightbox-img {
            width: 100%;
            height: 100%;
            display: block;
            object-fit: contain;
            cursor: pointer;
            transition: transform 0.3s;
        }

        .lightbox-img:hover {
            transform: scale(1.01);
        }

        .nav-controls {
            display: none;
        }

        .animate-up {
            opacity: 0;
            transform: translateY(30px);
            transition: opacity 0.8s ease-out, transform 0.8s ease-out;
        }
        
        .no-js .animate-up,
        .js-fallback .animate-up {
            opacity: 1;
            transform: none;
        }

        @media print {
            .nav-controls,
            .bg-element {
                display: none;
            }

            body {
                overflow: visible;
            }

            .slide {
                height: 100vh;
                page-break-after: always;
                padding: 2rem;
                background: #fff !important;
                color: #000 !important;
            }

            .bg-card {
                background: none !important;
                color: #000 !important;
                border-color: #eee !important;
            }
        }
    
        .hud-status {
            position: fixed; top: 1.5rem; left: 3.5rem;
            display: flex; align-items: center; gap: 12px;
            font-family: 'Roboto Mono', monospace; font-size: 0.75rem; color: var(--accent); z-index: 1000;
        }
        .hud-dot {
            width: 6px; height: 6px; background-color: #ff00ea; border-radius: 50%;
            animation: blinkHUD 1.5s infinite;
        }
        @keyframes blinkHUD { 0%, 100% { opacity: 1; box-shadow: 0 0 8px #ff00ea; } 50% { opacity: 0.3; box-shadow: none; } }
        .hud-progress-container {
            position: fixed; top: 0; left: 0; width: 100%; height: 2px; background: rgba(255,255,255,0.05); z-index: 1001;
        }
        .hud-progress-bar {
            height: 100%; background: var(--accent); width: 0%; transition: width 0.5s ease; box-shadow: 0 0 10px var(--accent);
        }
        .hud-corners {
            position: fixed; top: 0; left: 0; width: 100%; height: 100%; pointer-events: none; z-index: 999;
        }
        .corner { position: absolute; width: 20px; height: 20px; border: 1px solid var(--accent); opacity: 0.5; }
        .corner.topleft { top: 20px; left: 20px; border-right: none; border-bottom: none; }
        .corner.topright { top: 20px; right: 20px; border-left: none; border-bottom: none; }
        .corner.bottomleft { bottom: 20px; left: 20px; border-right: none; border-top: none; }
        .corner.bottomright { bottom: 20px; right: 20px; border-left: none; border-top: none; }

        .presenter-card {
            box-shadow: inset 0 1px 0 rgba(255,255,255,0.1), inset 1px 0 0 rgba(255,255,255,0.05), 0 8px 32px rgba(0,0,0,0.3);
            position: relative; overflow: hidden;
        }
        .presenter-card::before {
            content: ''; position: absolute; top: 0; left: 0; right: 0; bottom: 0;
            background-image: url("data:image/svg+xml,%3Csvg viewBox='0 0 200 200' xmlns='http://www.w3.org/2000/svg'%3E%3Cfilter id='n'%3E%3CfeTurbulence type='fractalNoise' baseFrequency='0.85' numOctaves='3' stitchTiles='stitch'/%3E%3C/filter%3E%3Crect width='100%25' height='100%25' filter='url(%23n)' opacity='0.05'/%3E%3C/svg%3E");
            pointer-events: none; z-index: -1; opacity: 0.15;
        }

        .mag-nav {
            position: fixed; top: 80px; height: calc(100vh - 80px); width: 8vw; z-index: 1000;
            display: flex; align-items: center; justify-content: center;
            cursor: pointer; opacity: 0; transition: opacity 0.4s, background 0.4s;
            background: rgba(0, 242, 255, 0);
        }
        .mag-nav.left { left: 0; border-left: 2px solid transparent; }
        .mag-nav.right { right: 0; border-right: 2px solid transparent; }
        .mag-nav:hover { opacity: 1; background: rgba(0, 242, 255, 0.02); }
        .mag-nav.left:hover { border-left-color: var(--accent); }
        .mag-nav.right:hover { border-right-color: var(--accent); }
        .mag-nav span {
            color: var(--accent); font-family: 'Roboto Mono', monospace; font-size: 0.85rem;
            transform: rotate(-90deg); letter-spacing: 0.3em; white-space: nowrap; user-select: none;
        }
        .mag-nav.right span { transform: rotate(90deg); }

        .sidebar {
            position: fixed; top: 0; left: -380px; width: 380px; height: 100vh;
            background: rgba(15, 23, 42, 0.95); backdrop-filter: blur(25px);
            border-right: 2px solid var(--accent);
            z-index: 3000; transition: left 0.4s cubic-bezier(0.25, 1, 0.5, 1);
            display: flex; flex-direction: column; padding: 2rem 1.5rem;
            box-shadow: 20px 0 50px rgba(0,0,0,0.5);
        }
        .sidebar.open { left: 0; }
        .sidebar-close {
            position: absolute; top: 1.5rem; right: 1.5rem; color: var(--text-dim);
            cursor: pointer; transition: color 0.3s;
        }
        .sidebar-close:hover { color: var(--accent); }
        .sidebar-header {
            font-family: 'Outfit', sans-serif; font-weight: 700; color: var(--text-main);
            font-size: 1.2rem; margin-bottom: 2rem; padding-bottom: 1rem; border-bottom: 1px solid var(--glass-border);
        }
        .sidebar-list { display: flex; flex-direction: column; gap: 0.5rem; overflow-y: auto; }
        .sidebar-item {
            padding: 0.75rem; border-radius: 8px; cursor: pointer; transition: all 0.3s;
            display: flex; flex-direction: column; gap: 0.5rem;
            border: 1px solid transparent; background: rgba(255,255,255,0.02);
            position: relative;
        }
        .sidebar-item:hover { background: var(--accent-soft); border-color: var(--accent); }
        .sidebar-item.active { background: rgba(0, 242, 255, 0.15); border-left: 4px solid var(--accent); }
        .thumb-container {
            width: 100%; aspect-ratio: 16/9; background: var(--bg-deep);
            border-radius: 4px; overflow: hidden; position: relative;
            border: 1px solid rgba(255,255,255,0.1);
        }
        .sidebar-item-header { display: flex; justify-content: space-between; align-items: center; }
        .sidebar-item-num { font-family: 'Roboto Mono', monospace; font-size: 0.75rem; color: var(--accent); }
        .sidebar-item-title { font-size: 0.9rem; color: var(--text-main); line-height: 1.3; }

        .menu-trigger {
            cursor: pointer; display: flex; align-items: center; justify-content: center;
            color: var(--text-dim); transition: color 0.3s;
        }
        .menu-trigger:hover { color: var(--accent); }

        .summary-item, .roadmap-item {
            margin-bottom: 1.5rem;
            padding: 1.2rem;
            background: rgba(255,255,255,0.03);
            border-radius: 0.8rem;
            border: 1px solid rgba(255,255,255,0.05);
            transition: all 0.3s ease;
            display: flex;
            align-items: flex-start;
            gap: 1.2rem;
        }
        .summary-item:hover, .roadmap-item:hover {
            background: rgba(255,255,255,0.08);
            transform: translateX(10px);
            border-color: var(--accent);
        }
        .summary-item i, .summary-item svg, .roadmap-item i, .roadmap-item svg {
            color: var(--accent);
            flex-shrink: 0;
            margin-top: 4px;
            width: 1.6rem !important;
            height: 1.6rem !important;
        }

        .roadmap-item i {
            animation: pulse-rocket 2s infinite ease-in-out;
        }

        @keyframes pulse-rocket {
            0% { transform: scale(1); filter: drop-shadow(0 0 5px var(--accent)); }
            50% { transform: scale(1.1); filter: drop-shadow(0 0 15px var(--accent)); }
            100% { transform: scale(1); filter: drop-shadow(0 0 5px var(--accent)); }
        }
        .rocket-glow {
            position: absolute;
            width: 200px;
            height: 200px;
            background: radial-gradient(circle, var(--accent-soft) 0%, transparent 70%);
            z-index: -1;
        }
    </style>
</head>

<body>

    <div class="bg-element bg-1"></div>
    <div class="bg-element bg-2"></div>

    <div class="hud-status">
        <div class="menu-trigger" id="menuTrigger" title="Открыть оглавление"><i data-lucide="menu" style="width: 18px; height: 18px;"></i></div>
        <div class="hud-dot" style="margin-left: 8px;"></div>
    </div>
    <div class="hud-progress-container"><div class="hud-progress-bar" id="hudProgress"></div></div>
    <div class="hud-corners">
        <div class="corner topleft"></div><div class="corner topright"></div>
        <div class="corner bottomleft"></div><div class="corner bottomright"></div>
    </div>
    <div class="mag-nav left" id="magPrev"><span>// ПРЕДЫДУЩИЙ</span></div>
    <div class="mag-nav right" id="magNext"><span>// СЛЕДУЮЩИЙ</span></div>

    <aside class="sidebar" id="sidebar">
        <div class="sidebar-close" id="sidebarClose"><i data-lucide="x"></i></div>
        <div class="sidebar-header"><i data-lucide="layers" style="width: 20px; vertical-align: middle; margin-right: 8px;"></i>Оглавление</div>
        <div class="sidebar-list" id="sidebarList">
        </div>
    </aside>

    <div class="presentation" id="presentation">
    <!-- SLIDES_START -->"""


BASE_HTML_TAIL = """
    <!-- SLIDES_END -->
    </div>

    <div class="lightbox" id="lightbox">
        <div class="lightbox-content">
            <img src="" alt="Full view" class="lightbox-img" id="lightboxImg">
        </div>
    </div>

    <script>
        document.documentElement.classList.add('js');
        
        if (typeof lucide !== 'undefined') {
            lucide.createIcons();
        }
        
        let currentSlide = 0;
        const presentation = document.getElementById('presentation');
        const totalSlides = presentation ? presentation.querySelectorAll('.slide').length : 0;

        function updateNavigation() {
            if (!presentation) return;
            presentation.style.transform = `translateX(-${currentSlide * 100}vw)`;
            document.querySelectorAll('.sidebar-item').forEach((item, index) => {
                if (index === currentSlide) item.classList.add('active');
                else item.classList.remove('active');
            });
            const progress = ((currentSlide + 1) / totalSlides) * 100;
            const hudBar = document.getElementById('hudProgress');
            if (hudBar) hudBar.style.width = progress + '%';
            const btnPrev = document.getElementById('magPrev');
            const btnNext = document.getElementById('magNext');
            if(btnPrev) btnPrev.style.display = currentSlide === 0 ? 'none' : 'flex';
            if(btnNext) btnNext.style.display = currentSlide === totalSlides - 1 ? 'none' : 'flex';

            const activeSlideItems = presentation.querySelectorAll('.slide')[currentSlide].querySelectorAll('.animate-up');
            if (activeSlideItems.length === 0) return;
            
            if (typeof gsap !== 'undefined') {
                const textItems = [];
                const visualItems = [];

                activeSlideItems.forEach(el => {
                    if (el.closest('.img-stack') || el.closest('.viz-card') || el.closest('.viz-item') || el.tagName === 'IMG') {
                        visualItems.push(el);
                    } else {
                        textItems.push(el);
                    }
                });

                if (visualItems.length > 0) {
                    gsap.set(visualItems, { opacity: 1, y: 0, scale: 1 });
                }

                if (textItems.length > 0) {
                    gsap.fromTo(textItems,
                        { opacity: 0, y: 30 },
                        { opacity: 1, y: 0, duration: 0.5, stagger: 0.1, ease: "power2.out" }
                    );
                }

                const activeMarkers = presentation.querySelectorAll('.slide')[currentSlide].querySelectorAll('.animate-marker');
                if (activeMarkers.length > 0) {
                    gsap.fromTo(activeMarkers,
                        { opacity: 0, scale: 0 },
                        { 
                            opacity: 1, 
                            scale: 1, 
                            duration: 0.4, 
                            stagger: 0.1, 
                            ease: "back.out(1.7)",
                            onComplete: function() {
                                this.targets().forEach(el => {
                                    const svg = el.tagName.toLowerCase() === 'svg' ? el : el.querySelector('svg');
                                    if (svg) svg.classList.add('marker-pulse');
                                    else el.classList.add('marker-pulse');
                                });
                            }
                        }
                    );
                }
            } else {
                activeSlideItems.forEach(el => {
                    el.style.opacity = '1';
                    el.style.transform = 'none';
                });
                document.body.classList.add('js-fallback');
            }
        }

        setTimeout(() => {
            if (typeof gsap === 'undefined' && !document.body.classList.contains('js-fallback')) {
                document.querySelectorAll('.animate-up').forEach(el => {
                    el.style.opacity = '1';
                    el.style.transform = 'none';
                });
                document.body.classList.add('js-fallback');
                console.log('[Fallback] GSAP not loaded, showing content');
            }
        }, 2000);

        const mP = document.getElementById('magPrev');
        const mN = document.getElementById('magNext');
        if(mP) mP.addEventListener('click', () => { if (currentSlide > 0) { currentSlide--; updateNavigation(); } });
        if(mN) mN.addEventListener('click', () => { if (currentSlide < totalSlides - 1) { currentSlide++; updateNavigation(); } });

        const slides = presentation ? presentation.querySelectorAll('.slide') : [];
        const sidebarList = document.getElementById('sidebarList');
        slides.forEach((slide, index) => {
            let titleEl = slide.querySelector('.slide-title');
            let headingEl = slide.querySelector('.main-heading');
            let titleText = "Слайд " + (index + 1);
            if (titleEl && titleEl.innerText) {
                titleText = titleEl.innerText;
            } else if (headingEl && headingEl.innerText) {
                titleText = "Главный титул";
            }
            const item = document.createElement('div');
            item.className = 'sidebar-item';
            item.innerHTML = `
                <div class="sidebar-item-header">
                    <span class="sidebar-item-num">0${index+1} / 0${slides.length}</span>
                    <span class="sidebar-item-title">${titleText}</span>
                </div>
                <div class="thumb-container" id="thumb-container-${index}"></div>
            `;
            item.addEventListener('click', () => {
                currentSlide = index;
                updateNavigation();
            });
            sidebarList.appendChild(item);

            setTimeout(() => {
                const tContainer = document.getElementById(`thumb-container-${index}`);
                if (!tContainer) return;
                const clone = slide.cloneNode(true);
                clone.removeAttribute('id');
                clone.querySelectorAll('[id]').forEach(el => el.removeAttribute('id'));
                clone.style.width = '1920px';
                clone.style.height = '1080px';
                clone.style.position = 'absolute';
                clone.style.top = '0';
                clone.style.left = '0';
                clone.style.display = 'flex';
                clone.style.padding = '4rem 8rem';
                const scale = tContainer.clientWidth / 1920;
                clone.style.transform = `scale(${scale})`;
                clone.style.transformOrigin = 'top left';
                clone.style.pointerEvents = 'none';
                clone.querySelectorAll('.animate-up').forEach(el => {
                    el.style.opacity = '1';
                    el.style.transform = 'none';
                    el.style.clipPath = 'none';
                });
                tContainer.appendChild(clone);
            }, 100);
        });

        const sidebar = document.getElementById('sidebar');
        const menuTrigger = document.getElementById('menuTrigger');
        const sidebarClose = document.getElementById('sidebarClose');

        if (menuTrigger && sidebar) menuTrigger.addEventListener('click', () => { sidebar.classList.add('open'); });
        if (sidebarClose && sidebar) sidebarClose.addEventListener('click', () => { sidebar.classList.remove('open'); });

        updateNavigation();
        
        (function initLucide() {
            if (typeof lucide !== 'undefined') {
                lucide.createIcons();
            } else {
                setTimeout(initLucide, 100);
            }
        })();

        document.addEventListener('keydown', (e) => {
            if (e.key === 'ArrowRight' && currentSlide < totalSlides - 1) { currentSlide++; updateNavigation(); }
            if (e.key === 'ArrowLeft' && currentSlide > 0) { currentSlide--; updateNavigation(); }
            if (e.key === 'Escape') closeLightbox();
        });

        const lightbox = document.getElementById('lightbox');
        const lightboxImg = document.getElementById('lightboxImg');

        if (lightbox && lightboxImg) {
            document.querySelectorAll('.viz-box').forEach(box => {
                box.addEventListener('click', () => {
                    const img = box.querySelector('img');
                    if (img) {
                        lightboxImg.src = img.src;
                        lightbox.style.display = 'flex';
                        if (typeof gsap !== 'undefined') {
                            gsap.to(lightbox, { opacity: 1, duration: 0.4, ease: "power2.out" });
                            gsap.fromTo('.lightbox-content',
                                { scale: 0.8, y: 50 },
                                { scale: 1, y: 0, duration: 0.5, ease: "back.out(1.7)" }
                            );
                        } else {
                            lightbox.style.opacity = 1;
                        }
                    }
                });
            });

            function closeLightbox() {
                if (lightbox.style.display === 'flex') {
                    if (typeof gsap !== 'undefined') {
                        gsap.to(lightbox, {
                            opacity: 0,
                            duration: 0.3,
                            onComplete: () => { lightbox.style.display = 'none'; }
                        });
                    } else {
                        lightbox.style.display = 'none';
                    }
                }
            }

            lightboxImg.addEventListener('click', closeLightbox);
            lightbox.addEventListener('click', (e) => {
                if (e.target === lightbox) closeLightbox();
            });
        }

        async function applyDynamicDensity() {
            const panels = document.querySelectorAll('.analytical-panel');
            
            for (const panel of panels) {
                panel.style.setProperty('--squeeze-factor', '1.0');
                panel.style.fontSize = "";
                panel.style.overflowY = 'hidden';
                panel.style.paddingBottom = '0px';

                await new Promise(r => requestAnimationFrame(r));
                
                let sH = panel.scrollHeight;
                let cH = panel.clientHeight;

                if (sH > cH + 1) {
                    let factor = (cH / sH) * 0.93;
                    factor = Math.max(0.60, factor);
                    
                    panel.style.setProperty('--squeeze-factor', factor.toFixed(3));
                    panel.style.paddingBottom = '4px';
                    
                    await new Promise(r => requestAnimationFrame(r));
                    
                    if (panel.scrollHeight > panel.clientHeight + 1) {
                        if (factor > 0.61) {
                            factor = Math.max(0.60, factor * 0.95);
                            panel.style.setProperty('--squeeze-factor', factor.toFixed(3));
                            await new Promise(r => requestAnimationFrame(r));
                        }
                        
                        if (panel.scrollHeight > panel.clientHeight + 1) {
                            panel.style.fontSize = "0.92em";
                            await new Promise(r => requestAnimationFrame(r));
                            
                            let finalSH = panel.scrollHeight;
                            let finalCH = panel.clientHeight;
                            if (finalSH > finalCH + 1) {
                                let finalFactor = Math.max(0.60, (finalCH / finalSH) * 0.95);
                                panel.style.setProperty('--squeeze-factor', finalFactor.toFixed(3));
                                await new Promise(r => requestAnimationFrame(r));
                            }
                        }
                    }

                    if (panel.scrollHeight > panel.clientHeight + 2) {
                        panel.style.overflowY = 'auto';
                    }
                }
            }
        }

        window.addEventListener('load', () => {
            setTimeout(applyDynamicDensity, 500);
        });
        
        let resizeTimer;
        window.addEventListener('resize', () => {
            clearTimeout(resizeTimer);
            resizeTimer = setTimeout(applyDynamicDensity, 200);
        });
    </script>
</body>
</html>
"""
