try:
    import streamlit as st
    import streamlit.components.v1 as components
except ImportError:
    st = None
    components = None

TAILWIND_CDN = '<script src="https://cdn.tailwindcss.com"></script>'

def render_tailwind_header(title: str, subtitle: str, badge_text: str = "Active Fiscal 2024", height: int = 120):
    """
    Renders an executive Tailwind-styled header banner inside an isolated HTML iframe.
    """
    html_content = f"""
    <!DOCTYPE html>
    <html>
    <head>
        {TAILWIND_CDN}
        <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
        <style>
            body {{ font-family: 'Inter', sans-serif; background: transparent; margin: 0; padding: 0; }}
        </style>
    </head>
    <body class="p-2">
        <div class="bg-gradient-to-r from-slate-900 via-blue-900 to-indigo-900 rounded-2xl p-5 shadow-lg border border-slate-700/50 flex flex-col md:flex-row justify-between items-start md:items-center">
            <div>
                <div class="flex items-center space-x-3 mb-1">
                    <span class="inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-semibold bg-emerald-500/20 text-emerald-300 border border-emerald-500/30">
                        {badge_text}
                    </span>
                    <span class="text-xs font-medium text-slate-400">ABES MCA Capstone | Project ID: P052 | ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01</span>
                </div>
                <h1 class="text-2xl font-bold tracking-tight text-white">{title}</h1>
                <p class="text-xs text-slate-300 mt-0.5">{subtitle}</p>
            </div>
            <div class="mt-3 md:mt-0 flex items-center space-x-2 text-xs text-slate-300">
                <span class="px-3 py-1.5 bg-slate-800/80 rounded-lg border border-slate-700 text-slate-200">
                    Analyst: <strong class="text-white">Anu Tyagi</strong>
                </span>
                <span class="px-3 py-1.5 bg-slate-800/80 rounded-lg border border-slate-700 text-slate-200">
                    Mentor: <strong class="text-white">Mr. Dinesh</strong>
                </span>
            </div>
        </div>
    </body>
    </html>
    """
    if components is not None:
        components.html(html_content, height=height, scrolling=False)
    return html_content


def render_kpi_cards(kpis: list, height: int = 150):
    """
    Renders a responsive multi-column grid of Tailwind-styled executive KPI cards.
    """
    card_htmls = []
    for k in kpis:
        title = k.get("title", "")
        value = k.get("value", "")
        delta = k.get("delta", "")
        positive = k.get("positive", True)
        
        delta_color = "text-emerald-700 bg-emerald-50 border-emerald-200" if positive else "text-rose-700 bg-rose-50 border-rose-200"
        arrow = "↑" if positive else "↓"
        
        card = f"""
        <div class="bg-white rounded-xl p-4 shadow-sm border border-slate-200 hover:shadow-md transition-shadow duration-200">
            <div class="text-xs font-semibold tracking-wide text-slate-500 uppercase">{title}</div>
            <div class="text-2xl font-extrabold text-slate-900 mt-1.5 tracking-tight">{value}</div>
            <div class="mt-2.5 flex items-center">
                <span class="inline-flex items-center px-2 py-0.5 rounded-md text-xs font-semibold border {delta_color}">
                    <span class="mr-1">{arrow}</span>{delta}
                </span>
                <span class="text-[11px] text-slate-400 ml-2">vs benchmark</span>
            </div>
        </div>
        """
        card_htmls.append(card)
        
    grid_cols = len(kpis) if len(kpis) <= 4 else 4
    html_content = f"""
    <!DOCTYPE html>
    <html>
    <head>
        {TAILWIND_CDN}
        <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">
        <style>
            body {{ font-family: 'Inter', sans-serif; background: transparent; margin: 0; padding: 0; }}
        </style>
    </head>
    <body class="p-2">
        <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-{grid_cols} gap-4">
            {''.join(card_htmls)}
        </div>
    </body>
    </html>
    """
    if components is not None:
        components.html(html_content, height=height, scrolling=False)
    return html_content


def render_insight_box(title: str, message: str, level: str = "info", height: int = 110):
    """
    Renders an analytical diagnostic callout box.
    """
    themes = {
        "info": ("bg-blue-50 border-blue-200 text-blue-900", "text-blue-700", "INFO"),
        "warning": ("bg-amber-50 border-amber-200 text-amber-900", "text-amber-700", "WARNING"),
        "danger": ("bg-rose-50 border-rose-200 text-rose-900", "text-rose-700", "ALERT"),
        "success": ("bg-emerald-50 border-emerald-200 text-emerald-900", "text-emerald-700", "INSIGHT")
    }
    bg_border_text, badge_color, badge_text = themes.get(level, themes["info"])
    
    html_content = f"""
    <!DOCTYPE html>
    <html>
    <head>
        {TAILWIND_CDN}
        <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
        <style>
            body {{ font-family: 'Inter', sans-serif; background: transparent; margin: 0; padding: 0; }}
        </style>
    </head>
    <body class="p-1">
        <div class="rounded-xl p-3.5 border shadow-sm {bg_border_text}">
            <div class="flex items-center space-x-2">
                <span class="px-2 py-0.5 rounded text-[10px] font-bold tracking-wider uppercase bg-white/80 border border-current {badge_color}">
                    {badge_text}
                </span>
                <h4 class="text-sm font-bold">{title}</h4>
            </div>
            <p class="text-xs mt-1.5 leading-relaxed opacity-90">{message}</p>
        </div>
    </body>
    </html>
    """
    if components is not None:
        components.html(html_content, height=height, scrolling=False)
    return html_content
