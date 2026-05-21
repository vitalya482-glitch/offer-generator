from __future__ import annotations


def ui_scale(width: int, height: int) -> float:
    width = max(width, 900)
    height = max(height, 620)
    return max(0.86, min(1.25, min(width / 1440, height / 900)))


def stylesheet(scale: float) -> str:
    def px(value: int) -> int:
        return max(1, int(value * scale))

    return f"""
        QMainWindow {{ background: #F4F6F8; }}
        #ContentScroll {{ background: #F4F6F8; border: none; }}
        #Sidebar {{ background: #15171B; color: #FFFFFF; }}
        #Content {{ background: #F4F6F8; }}
        #Brand {{ color: #FFFFFF; font-size: {px(34)}px; font-weight: 900; letter-spacing: 2px; }}
        #SideTitle {{ color: #FFFFFF; font-size: {px(23)}px; font-weight: 700; }}
        #SideSubtitle {{ color: #B8C0CC; font-size: {px(13)}px; line-height: 1.4; }}
        #Badge {{ background: #D71920; color: white; border-radius: {px(16)}px; padding: {px(8)}px {px(12)}px; font-weight: 700; }}
        #PageTitle {{ color: #171A1F; font-size: {px(28)}px; font-weight: 800; }}
        #PageSubtitle {{ color: #667085; font-size: {px(13)}px; }}
        #Card {{ background: #FFFFFF; border: 1px solid #E7EAF0; border-radius: {px(18)}px; }}
        #CardTitle {{ color: #171A1F; font-size: {px(15)}px; font-weight: 800; }}
        #FormLabel {{ color: #344054; font-weight: 700; font-size: {px(12)}px; }}
        QLineEdit, QComboBox, QTextEdit {{
            background-color: #FFFFFF;
            color: #101828;
            border: 2px solid #D0D5DD;
            border-radius: {px(12)}px;
            padding: 0 {px(12)}px;
            min-height: {px(38)}px;
            font-size: {px(13)}px;
            selection-background-color: #D71920;
        }}
        QLineEdit:focus, QComboBox:focus, QTextEdit:focus {{ border: 2px solid #D71920; }}
        QTextEdit {{ min-height: {px(150)}px; padding: {px(12)}px {px(14)}px;}}
        QPushButton {{ border-radius: {px(11)}px; padding: {px(9)}px {px(14)}px; font-weight: 800; font-size: {px(12)}px; }}
        #PrimaryButton {{ background: #D71920; color: white; border: 1px solid #D71920; }}
        #PrimaryButton:hover {{ background: #B9151B; }}
        #GhostButton {{ background: #FFFFFF; color: #1D2939; border: 1px solid #D0D5DD; }}
        #GhostButton:hover {{ border: 1px solid #D71920; color: #D71920; }}
        QLabel {{ color: #475467; font-size: {px(12)}px; }}
    """
