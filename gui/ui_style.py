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
        #SidebarScroll {{ background: #15171B; border: none; }}
        #Sidebar {{ background: #15171B; color: #FFFFFF; }}
        #Content {{ background: #F4F6F8; }}
        #Brand {{ color: #FFFFFF; font-size: {px(34)}px; font-weight: 900; letter-spacing: 2px; }}
        #SideTitle {{ color: #FFFFFF; font-size: {px(23)}px; font-weight: 700; }}
        #SideSubtitle {{ color: #B8C0CC; font-size: {px(11)}px; line-height: 1.25; }}
        #SidebarSectionTitle {{ color: #FFFFFF; font-size: {px(13)}px; font-weight: 800; margin-top: {px(2)}px; }}
        #SidebarFormLabel {{ color: #B8C0CC; font-size: {px(10)}px; font-weight: 700; }}
        #SidebarHint {{ color: #8A94A6; font-size: {px(9)}px; }}
        #SidebarFooter {{ color: #667085; font-size: {px(9)}px; }}
        #SidebarInput {{
            background-color: #20242B;
            color: #FFFFFF;
            border: 1px solid #343A46;
            border-radius: {px(8)}px;
            padding: 0 {px(8)}px;
            min-height: {px(26)}px;
            max-height: {px(28)}px;
            font-size: {px(10)}px;
        }}
        #SidebarInput:focus {{ border: 1px solid #D71920; }}
        #SidebarButton {{ background: #D71920; color: white; border: 1px solid #D71920; border-radius: {px(9)}px; min-height: {px(28)}px; max-height: {px(30)}px; font-size: {px(10)}px; }}
        #SidebarButton:hover {{ background: #B9151B; }}
        #Badge {{ background: #D71920; color: white; border-radius: {px(14)}px; padding: {px(6)}px {px(10)}px; font-weight: 700; font-size: {px(10)}px; }}

        QRadioButton#SidebarRadio {
            color: #E5E7EB;
            font-size: {px(10)}px;
            spacing: {px(5)}px;
            padding: {px(1)}px 0;
        }
        QRadioButton#SidebarRadio:disabled {
            color: #98A2B3;
        }
        QRadioButton#SidebarRadio::indicator {
            width: {px(13)}px;
            height: {px(13)}px;
        }
        QRadioButton#SidebarRadio::indicator:unchecked {
            border: 1px solid #4B5563;
            border-radius: {px(7)}px;
            background: #15171B;
        }
        QRadioButton#SidebarRadio::indicator:checked {
            border: 1px solid #2E90FA;
            border-radius: {px(7)}px;
            background: #2E90FA;
        }
        QScrollBar:vertical {
            background: transparent;
            width: {px(7)}px;
            margin: 0;
        }
        QScrollBar::handle:vertical {
            background: #3B4250;
            border-radius: {px(3)}px;
            min-height: {px(24)}px;
        }
        QScrollBar::add-line:vertical,
        QScrollBar::sub-line:vertical {
            height: 0;
            background: transparent;
        }
        QScrollBar::add-page:vertical,
        QScrollBar::sub-page:vertical {
            background: transparent;
        }
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
