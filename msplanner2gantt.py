import os
import sys
import tkinter as tk
from tkinter import filedialog, simpledialog
from datetime import datetime
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


def select_planner_file():
    """Opens a file dialog for the user to choose their MS Planner export."""
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    file_path = filedialog.askopenfilename(
        title="Select MS Planner Export Spreadsheet",
        filetypes=[("Excel Files", "*.xlsx *.xls"), ("All Files", "*.*")]
    )
    root.destroy()
    return file_path


def prompt_project_start_date(default_date_str=""):
    """Pops up an input dialog prompting the user for the actual project start date[cite: 4, 5]."""
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    val = simpledialog.askstring(
        "Project Start Date",
        f"Enter Actual Project Start Date (YYYY-MM or YYYY-MM-DD):\n"
        f"(Leave blank to default to earliest date: {default_date_str})",
        parent=root
    )
    root.destroy()
    return val


def parse_checklist(raw_checklist, completed_count_str):
    """
    Splits semicolon-separated checklist items and tags each item as
    completed [✓] or pending [ ] based on the 'Completed Checklist Items' count[cite: 4, 5].
    """
    if pd.isna(raw_checklist) or str(raw_checklist).strip() == "" or str(raw_checklist).lower() == "nan":
        return "No checklist items"

    items = [item.strip() for item in str(raw_checklist).split(';') if item.strip()]
    if not items:
        return "No checklist items"

    num_completed = 0
    if pd.notna(completed_count_str) and "/" in str(completed_count_str):
        try:
            num_completed = int(str(completed_count_str).split('/')[0].strip())
        except ValueError:
            num_completed = 0
    elif pd.notna(completed_count_str):
        try:
            num_completed = int(completed_count_str)
        except ValueError:
            num_completed = 0

    formatted_lines = []
    for i, itm in enumerate(items):
        if i < num_completed:
            formatted_lines.append(f"[✓] {itm}")
        else:
            formatted_lines.append(f"[ ] {itm}")

    return "\n".join(formatted_lines)


def generate_gantt_from_planner(planner_excel_path=None, project_start_date=None, output_excel_path=None):
    """
    Builds a custom Gantt chart from MS Planner exports:
      1. Completed tasks positioned on top.
      2. Dynamic project start date parameter (red vertical line; defaults to earliest date if blank)[cite: 4, 5].
      3. In-progress tasks: Blue for past/current months; Gray for future months (after today's month).
      4. Blank 'Expense' column inserted between Due Date and Days (formatted with currency ready for input).
      5. Checklist items tagged as [✓] done or [ ] pending.
      6. Output file saved in the same directory with a '_YYYY-MM-DD' date suffix[cite: 4, 5].
    """
    # 1. Prompt file selection if not provided
    if not planner_excel_path:
        planner_excel_path = select_planner_file()
        if not planner_excel_path:
            print("No file selected. Operation cancelled.")
            return

    # 2. Read Planner export
    try:
        df = pd.read_excel(planner_excel_path, sheet_name="Consolidated Data")
    except Exception:
        df = pd.read_excel(planner_excel_path)

    df = df.dropna(subset=['Task Name']).copy()
    df['Start date'] = pd.to_datetime(df['Start date'], errors='coerce')
    df['Due date'] = pd.to_datetime(df['Due date'], errors='coerce')

    today = pd.Timestamp.now().normalize()
    df['Start date'] = df['Start date'].fillna(today)
    df['Due date'] = df['Due date'].fillna(df['Start date'] + pd.Timedelta(days=30))
    df['Duration_Days'] = (df['Due date'] - df['Start date']).dt.days.clip(lower=1)

    # Sort: Completed tasks placed at the top, then sorted by Start date[cite: 4, 5]
    df['Status_Rank'] = df['Status'].apply(lambda x: 0 if str(x).strip().lower() == 'completed' else 1)
    df = df.sort_values(by=['Status_Rank', 'Start date']).reset_index(drop=True)

    earliest_task_date = df['Start date'].min()

    # Dynamic Project Start Date Parameter
    if project_start_date is None:
        user_input = prompt_project_start_date(default_date_str=earliest_task_date.strftime('%Y-%m'))
        if user_input and user_input.strip():
            parsed_dt = pd.to_datetime(user_input.strip(), errors='coerce')
            actual_start_dt = earliest_task_date if pd.isna(parsed_dt) else parsed_dt
        else:
            actual_start_dt = earliest_task_date
    else:
        if str(project_start_date).strip() == "":
            actual_start_dt = earliest_task_date
        else:
            parsed_dt = pd.to_datetime(str(project_start_date).strip(), errors='coerce')
            actual_start_dt = earliest_task_date if pd.isna(parsed_dt) else parsed_dt

    # Output file in the same folder with '_YYYY-MM-DD' suffix[cite: 4, 5]
    today_str = datetime.today().strftime('%Y-%m-%d')
    if not output_excel_path:
        folder_dir = os.path.dirname(os.path.abspath(planner_excel_path))
        base_name = os.path.splitext(os.path.basename(planner_excel_path))[0]
        output_excel_path = os.path.join(folder_dir, f"{base_name}_Gantt_Chart_{today_str}.xlsx")

    # Timeline bounds
    min_date = min(df['Start date'].min(), actual_start_dt, today)
    max_date = max(df['Due date'].max(), today)

    cur_y, cur_m = min_date.year, min_date.month
    end_y, end_m = max_date.year, max_date.month

    timeline_months = []
    while (cur_y < end_y) or (cur_y == end_y and cur_m <= end_m):
        timeline_months.append((cur_y, cur_m))
        cur_m += 1
        if cur_m > 12:
            cur_m = 1
            cur_y += 1

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    ws = wb.create_sheet(title="Gantt Chart & Status")
    ws.views.sheetView[0].showGridLines = True

    # Typography & Styles
    FONT_NAME = "Segoe UI"
    header_font = Font(name=FONT_NAME, size=10, bold=True, color="FFFFFF")
    title_font = Font(name=FONT_NAME, size=15, bold=True, color="1B365D")
    sub_font = Font(name=FONT_NAME, size=9, italic=True, color="555555")
    regular_font = Font(name=FONT_NAME, size=9, color="2C3E50")
    bold_font = Font(name=FONT_NAME, size=9, bold=True, color="2C3E50")

    navy_fill = PatternFill(start_color="1B365D", fill_type="solid")
    slate_fill = PatternFill(start_color="34495E", fill_type="solid")
    zebra_fill = PatternFill(start_color="F9FAFB", fill_type="solid")

    # Bar fills
    bar_completed = PatternFill(start_color="28A745", fill_type="solid")
    bar_active_current = PatternFill(start_color="3B82F6", fill_type="solid")
    bar_active_future = PatternFill(start_color="94A3B8", fill_type="solid")

    thin_gray = Side(style='thin', color='E5E7EB')
    thin_border = Border(left=thin_gray, right=thin_gray, top=thin_gray, bottom=thin_gray)

    # Solid red vertical line for Project Start Date[cite: 3, 4]
    red_solid_side = Side(style='medium', color='FF0000')

    status_palette = {
        'Completed': (PatternFill(start_color="D4EDDA", fill_type="solid"), Font(name=FONT_NAME, size=9, bold=True, color="155724")),
        'In progress': (PatternFill(start_color="DBEAFE", fill_type="solid"), Font(name=FONT_NAME, size=9, bold=True, color="1E40AF")),
        'Not started': (PatternFill(start_color="F3F4F6", fill_type="solid"), Font(name=FONT_NAME, size=9, bold=True, color="4B5563"))
    }

    priority_palette = {
        'Urgent': (PatternFill(start_color="FEE2E2", fill_type="solid"), Font(name=FONT_NAME, size=9, bold=True, color="991B1B")),
        'Important': (PatternFill(start_color="FED7AA", fill_type="solid"), Font(name=FONT_NAME, size=9, bold=True, color="9A3412")),
        'Medium': (PatternFill(start_color="FEF08A", fill_type="solid"), Font(name=FONT_NAME, size=9, bold=True, color="854D0E")),
        'Low': (PatternFill(start_color="E0F2FE", fill_type="solid"), Font(name=FONT_NAME, size=9, bold=True, color="075985"))
    }

    ws["B2"] = "Project Gantt Chart & Status Tracker"
    ws["B2"].font = title_font
    ws["B3"] = (
        f"Export Date: {today_str} | "
        f"Project Start: {actual_start_dt.strftime('%b %Y')} (Red Line) | "
        f"Active Months: Blue | Future Planned Months: Gray"
    )
    ws["B3"].font = sub_font

    # Metadata columns with 'Expense' inserted between 'Due Date' and 'Days'
    meta_headers = [
        ("Task ID", 9),
        ("Task Name", 34),
        ("Bucket / Stage", 20),
        ("Status", 13),
        ("Priority", 11),
        ("Assigned To", 15),
        ("Start Date", 12),
        ("Due Date", 12),
        ("Expense", 13),  # <-- Column J (left blank for manual input)
        ("Days", 8),       # <-- Column K
        ("Checklist Items & Status ([✓] Done, [ ] Pending)", 48)  # <-- Column L
    ]

    for c_idx, (title, width) in enumerate(meta_headers, start=2):
        cell = ws.cell(row=5, column=c_idx, value=title)
        cell.fill = navy_fill
        cell.font = header_font
        cell.border = thin_border
        cell.alignment = Alignment(horizontal="center" if c_idx not in (3, 4, 12) else "left", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(c_idx)].width = width

    ws.column_dimensions['A'].width = 3
    tl_start_col = 13  # Shifted to Column M because of the new Expense column

    start_col = None
    today_col = None

    year_map = {}
    for m_i, (yr, mo) in enumerate(timeline_months):
        c_num = tl_start_col + m_i
        year_map.setdefault(yr, []).append(c_num)

        if yr == actual_start_dt.year and mo == actual_start_dt.month:
            start_col = c_num
        if yr == today.year and mo == today.month:
            today_col = c_num

        m_cell = ws.cell(row=5, column=c_num, value=datetime(yr, mo, 1).strftime("%b '%y"))
        m_cell.fill = slate_fill
        m_cell.font = Font(name=FONT_NAME, size=8, bold=True, color="FFFFFF")

        # Header accent for Start and Current Month
        b_left = thin_gray
        if c_num == start_col:
            b_left = red_solid_side
            m_cell.fill = PatternFill(start_color="FFE5E5", fill_type="solid")
            m_cell.font = Font(name=FONT_NAME, size=8, bold=True, color="C00000")
        elif c_num == today_col:
            m_cell.fill = PatternFill(start_color="E2E8F0", fill_type="solid")
            m_cell.font = Font(name=FONT_NAME, size=8, bold=True, color="1E293B")

        m_cell.border = Border(left=b_left, right=thin_gray, top=thin_gray, bottom=thin_gray)
        m_cell.alignment = Alignment(horizontal="center", vertical="center", textRotation=90)
        ws.column_dimensions[get_column_letter(c_num)].width = 5.5

    # Year header row
    for yr, cols in year_map.items():
        ws.merge_cells(start_row=4, start_column=cols[0], end_row=4, end_column=cols[-1])
        y_cell = ws.cell(row=4, column=cols[0], value=str(yr))
        y_cell.fill = navy_fill
        y_cell.font = Font(name=FONT_NAME, size=10, bold=True, color="FFFFFF")
        y_cell.alignment = Alignment(horizontal="center", vertical="center")
        for c in cols:
            b_left = red_solid_side if (c == start_col) else thin_gray
            ws.cell(row=4, column=c).border = Border(left=b_left, right=thin_gray, top=thin_gray, bottom=thin_gray)

    # Top indicator labels above header
    if start_col:
        lbl = ws.cell(row=3, column=start_col, value=f"▼ Project Start ({actual_start_dt.strftime('%b %Y')})")
        lbl.font = Font(name=FONT_NAME, size=8.5, bold=True, color="C00000")
        lbl.alignment = Alignment(horizontal="left", vertical="center")
    if today_col:
        lbl2 = ws.cell(row=3, column=today_col, value=f"▼ Current Month ({today.strftime('%b %Y')})")
        lbl2.font = Font(name=FONT_NAME, size=8.5, bold=True, color="334155")
        lbl2.alignment = Alignment(horizontal="left", vertical="center")

    # Tasks rows
    start_r = 6
    for idx, row in df.iterrows():
        r = start_r + idx

        ws.cell(row=r, column=2, value=f"T-{idx+1:02d}").alignment = Alignment(horizontal="center", vertical="top")
        ws.cell(row=r, column=3, value=str(row.get('Task Name', ''))).alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        ws.cell(row=r, column=4, value=str(row.get('Bucket', ''))).alignment = Alignment(horizontal="left", vertical="top")

        st = str(row.get('Status', ''))
        c_st = ws.cell(row=r, column=5, value=st)
        c_st.alignment = Alignment(horizontal="center", vertical="top")
        if st in status_palette:
            c_st.fill, c_st.font = status_palette[st]

        pr = str(row.get('Priority', ''))
        c_pr = ws.cell(row=r, column=6, value=pr)
        c_pr.alignment = Alignment(horizontal="center", vertical="top")
        if pr in priority_palette:
            c_pr.fill, c_pr.font = priority_palette[pr]

        ws.cell(row=r, column=7, value=str(row.get('Assigned To', ''))).alignment = Alignment(horizontal="center", vertical="top")
        ws.cell(row=r, column=8, value=row['Start date'].strftime('%Y-%m-%d')).alignment = Alignment(horizontal="center", vertical="top")
        ws.cell(row=r, column=9, value=row['Due date'].strftime('%Y-%m-%d')).alignment = Alignment(horizontal="center", vertical="top")

        # Column 10: Blank 'Expense' cell with accounting currency formatting ($#,##0.00)
        c_exp = ws.cell(row=r, column=10, value="")
        c_exp.alignment = Alignment(horizontal="right", vertical="top")
        c_exp.number_format = '$#,##0.00'

        # Column 11: Duration in Days
        ws.cell(row=r, column=11, value=int(row['Duration_Days'])).alignment = Alignment(horizontal="center", vertical="top")

        # Column 12: Checklist Items
        raw_chk = row.get('Checklist Items', '')
        chk_count = row.get('Completed Checklist Items', '')
        formatted_chk = parse_checklist(raw_chk, chk_count)

        c_chk = ws.cell(row=r, column=12, value=formatted_chk)
        c_chk.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        c_chk.font = Font(name=FONT_NAME, size=8.5, color="2C3E50")

        for c in range(2, 13):
            cell = ws.cell(row=r, column=c)
            cell.border = thin_border
            if c not in (5, 6, 12):
                cell.font = regular_font
            if idx % 2 == 1 and c not in (5, 6):
                cell.fill = zebra_fill

        # Monthly timeline bars:
        t_start = row['Start date'].date()
        t_due = row['Due date'].date()

        for m_i, (yr, mo) in enumerate(timeline_months):
            c_num = tl_start_col + m_i
            m_s = datetime(yr, mo, 1).date()
            n_yr = yr if mo < 12 else yr + 1
            n_mo = mo + 1 if mo < 12 else 1
            m_e = (datetime(n_yr, n_mo, 1) - pd.Timedelta(days=1)).date()

            bar_cell = ws.cell(row=r, column=c_num)

            # Left border: Project Start Date in solid red
            b_l = red_solid_side if c_num == start_col else thin_gray
            bar_cell.border = Border(left=b_l, right=thin_gray, top=thin_gray, bottom=thin_gray)

            if not (t_due < m_s or t_start > m_e):
                if st == "Completed":
                    bar_cell.fill = bar_completed
                else:
                    # Blue for past/current months, Slate Gray for future months
                    is_future_month = (yr > today.year) or (yr == today.year and mo > today.month)
                    if is_future_month:
                        bar_cell.fill = bar_active_future
                    else:
                        bar_cell.fill = bar_active_current

    # Legend & Key
    leg_r = start_r + len(df) + 2
    ws.cell(row=leg_r, column=2, value="Legend:").font = bold_font

    ws.cell(row=leg_r, column=3, value="■ Completed Task").font = Font(name=FONT_NAME, size=9, bold=True, color="155724")
    ws.cell(row=leg_r, column=3).fill = PatternFill(start_color="D4EDDA", fill_type="solid")
    ws.cell(row=leg_r, column=3).alignment = Alignment(horizontal="center")

    ws.cell(row=leg_r, column=4, value="■ Active (Past / Current)").font = Font(name=FONT_NAME, size=9, bold=True, color="1E40AF")
    ws.cell(row=leg_r, column=4).fill = PatternFill(start_color="DBEAFE", fill_type="solid")
    ws.cell(row=leg_r, column=4).alignment = Alignment(horizontal="center")

    ws.cell(row=leg_r, column=5, value="■ Future Planned Work").font = Font(name=FONT_NAME, size=9, bold=True, color="475569")
    ws.cell(row=leg_r, column=5).fill = PatternFill(start_color="CBD5E1", fill_type="solid")
    ws.cell(row=leg_r, column=5).alignment = Alignment(horizontal="center")

    l_start = ws.cell(row=leg_r, column=6, value=f"| Project Start ({actual_start_dt.strftime('%b %Y')})")
    l_start.font = Font(name=FONT_NAME, size=9, bold=True, color="C00000")
    l_start.alignment = Alignment(horizontal="center")

    ws.cell(row=leg_r + 1, column=2, value="Checklist:").font = bold_font
    ws.cell(row=leg_r + 1, column=3, value="[✓] Completed Step").font = Font(name=FONT_NAME, size=9, color="155724")
    ws.cell(row=leg_r + 1, column=4, value="[ ] Pending Step").font = Font(name=FONT_NAME, size=9, color="7F8C8D")

    # Freeze panes at Column M (Timeline starts at Column M)
    ws.freeze_panes = "M6"
    wb.save(output_excel_path)
    print(f"\nSuccess! Gantt chart created at:\n{output_excel_path}")


if __name__ == "__main__":
    generate_gantt_from_planner()