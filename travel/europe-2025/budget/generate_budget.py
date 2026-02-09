"""Generate the Europe Trip 2025 budget Excel file with hyperlinks and formulas."""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, numbers
from openpyxl.utils import get_column_letter


def create_budget():
    wb = openpyxl.Workbook()

    # Styles
    header_font = Font(name="Calibri", bold=True, size=12, color="FFFFFF")
    header_fill = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
    subheader_font = Font(name="Calibri", bold=True, size=11, color="2F5496")
    subheader_fill = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")
    currency_format = '$#,##0.00'
    total_font = Font(name="Calibri", bold=True, size=11)
    total_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    grand_total_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    warning_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    link_font = Font(name="Calibri", size=10, color="0563C1", underline="single")
    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )

    def style_header_row(ws, row, max_col):
        for col in range(1, max_col + 1):
            cell = ws.cell(row=row, column=col)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", wrap_text=True)
            cell.border = thin_border

    def style_data_cell(ws, row, col, is_currency=False):
        cell = ws.cell(row=row, column=col)
        cell.border = thin_border
        cell.alignment = Alignment(wrap_text=True)
        if is_currency:
            cell.number_format = currency_format

    # ========================================
    # Sheet 1: Summary
    # ========================================
    ws_summary = wb.active
    ws_summary.title = "Summary"
    ws_summary.sheet_properties.tabColor = "2F5496"

    # Title
    ws_summary["A1"] = "Europe Trip 2025 - Budget Summary"
    ws_summary["A1"].font = Font(name="Calibri", bold=True, size=16, color="2F5496")
    ws_summary.merge_cells("A1:F1")

    ws_summary["A2"] = "4 Travelers | 21 Days | September 2025 | Toronto -> Lisbon -> Porto -> Barcelona -> Zurich -> Toronto"
    ws_summary["A2"].font = Font(name="Calibri", size=10, italic=True)
    ws_summary.merge_cells("A2:F2")

    ws_summary["A3"] = f"Total Budget: $10,000 USD"
    ws_summary["A3"].font = Font(name="Calibri", bold=True, size=12, color="006100")
    ws_summary.merge_cells("A3:F3")

    # Headers
    headers = ["Category", "Subcategory", "Cost (4 People)", "% of Budget", "Sheet Link", "Notes"]
    for col, h in enumerate(headers, 1):
        ws_summary.cell(row=5, column=col, value=h)
    style_header_row(ws_summary, 5, len(headers))

    # Data rows
    summary_data = [
        ("Flights", "All direct flights", 3840, "Flights", "TAP, Air Canada, SWISS, Vueling"),
        ("", "Toronto -> Lisbon", 1800, "", "TAP/Air Canada/Air Transat direct 7h"),
        ("", "Lisbon -> Barcelona", 200, "", "TAP/Vueling/Ryanair direct 2h"),
        ("", "Barcelona -> Zurich", 240, "", "SWISS/Vueling direct 1h45m"),
        ("", "Zurich -> Toronto", 1600, "", "Air Canada/SWISS direct 9h"),
        ("Accommodation", "20 nights total", 2550, "Accommodation", "3 hotel check-ins total"),
        ("", "Lisbon - Hotel da Baixa (5 nights)", 550, "", "~$110/night, family rooms"),
        ("", "Porto - Moov Hotel (2 nights)", 200, "", "~$100/night, superior rooms"),
        ("", "Barcelona - Aparthotel Arai (7 nights)", 840, "", "~$120/night, kitchenettes"),
        ("", "Zurich - Hotel Glockenhof (6 nights)", 960, "", "~$160/night, family suites"),
        ("Transportation", "Trains & local transit", 1800, "Transportation", "Swiss Pass + local transit"),
        ("", "Swiss Travel Pass (8-day x 4)", 1520, "", "Unlimited Swiss trains/buses/boats"),
        ("", "Portugal trains/metro", 140, "", "Lisbon-Porto Alfa Pendular + local"),
        ("", "Spain trains/metro", 140, "", "Barcelona day trips + T-Casual card"),
        ("Activities", "Nature-focused", 400, "Activities", "Entry fees, cable cars, boats"),
        ("", "Mountain cables/funiculars", 200, "", "Grindelwald First, Tibidabo, etc."),
        ("", "Park/waterfall entrances", 100, "", "Trummelbach Falls, etc."),
        ("", "Boat rides", 100, "", "Rhine Falls boat, misc."),
    ]

    row = 6
    category_total_rows = {}  # track total rows for each category
    for data in summary_data:
        cat, subcat, cost, sheet_ref, notes = data
        ws_summary.cell(row=row, column=1, value=cat)
        ws_summary.cell(row=row, column=2, value=subcat)
        ws_summary.cell(row=row, column=3, value=cost)
        ws_summary.cell(row=row, column=4).value = cost / 10000
        ws_summary.cell(row=row, column=4).number_format = "0.0%"
        if sheet_ref:
            ws_summary.cell(row=row, column=5, value=f"See '{sheet_ref}' sheet")
        ws_summary.cell(row=row, column=6, value=notes)

        for col in range(1, 7):
            style_data_cell(ws_summary, row, col, is_currency=(col == 3))

        if cat:  # Category header row
            ws_summary.cell(row=row, column=1).font = subheader_font
            ws_summary.cell(row=row, column=2).font = subheader_font
            for c in range(1, 7):
                ws_summary.cell(row=row, column=c).fill = subheader_fill

        row += 1

    # Grand Total
    row += 1
    ws_summary.cell(row=row, column=1, value="GRAND TOTAL")
    ws_summary.cell(row=row, column=1).font = Font(name="Calibri", bold=True, size=12)
    ws_summary.cell(row=row, column=3, value=8590)
    ws_summary.cell(row=row, column=3).number_format = currency_format
    ws_summary.cell(row=row, column=3).font = Font(name="Calibri", bold=True, size=12)
    ws_summary.cell(row=row, column=4).value = 8590 / 10000
    ws_summary.cell(row=row, column=4).number_format = "0.0%"
    for c in range(1, 7):
        ws_summary.cell(row=row, column=c).fill = grand_total_fill
        ws_summary.cell(row=row, column=c).border = thin_border

    # Remaining Buffer
    row += 1
    ws_summary.cell(row=row, column=1, value="REMAINING BUFFER")
    ws_summary.cell(row=row, column=1).font = Font(name="Calibri", bold=True, size=12, color="006100")
    ws_summary.cell(row=row, column=3).value = 1410
    ws_summary.cell(row=row, column=3).number_format = currency_format
    ws_summary.cell(row=row, column=3).font = Font(name="Calibri", bold=True, size=12, color="006100")
    ws_summary.cell(row=row, column=4).value = 1410 / 10000
    ws_summary.cell(row=row, column=4).number_format = "0.0%"
    ws_summary.cell(row=row, column=6, value="For food, emergencies, extras")
    for c in range(1, 7):
        ws_summary.cell(row=row, column=c).fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        ws_summary.cell(row=row, column=c).border = thin_border

    # Column widths
    ws_summary.column_dimensions["A"].width = 18
    ws_summary.column_dimensions["B"].width = 38
    ws_summary.column_dimensions["C"].width = 18
    ws_summary.column_dimensions["D"].width = 14
    ws_summary.column_dimensions["E"].width = 22
    ws_summary.column_dimensions["F"].width = 42

    # ========================================
    # Sheet 2: Flights
    # ========================================
    ws_flights = wb.create_sheet("Flights")
    ws_flights.sheet_properties.tabColor = "4472C4"

    ws_flights["A1"] = "Flights Budget - All Direct, No Layovers"
    ws_flights["A1"].font = Font(name="Calibri", bold=True, size=14, color="2F5496")
    ws_flights.merge_cells("A1:H1")

    flight_headers = ["Route", "Airline Options", "Duration", "Price/Person", "Qty", "Total (4 ppl)", "Frequency", "Source"]
    for col, h in enumerate(flight_headers, 1):
        ws_flights.cell(row=3, column=col, value=h)
    style_header_row(ws_flights, 3, len(flight_headers))

    flight_data = [
        ("Toronto -> Lisbon", "TAP Air Portugal / Air Canada / Air Transat", "7h direct", 450, 4, "=D4*E4",
         "TAP 8/wk, AC 4/wk, Transat seasonal", "https://www.flytap.com"),
        ("Lisbon -> Barcelona", "TAP / Vueling / Ryanair / easyJet", "2h direct", 50, 4, "=D5*E5",
         "12+ direct flights daily", "https://www.vueling.com"),
        ("Barcelona -> Zurich", "SWISS / Vueling", "1h 45m direct", 60, 4, "=D6*E6",
         "SWISS 21/wk, Vueling 12/wk", "https://www.swiss.com"),
        ("Zurich -> Toronto", "Air Canada / SWISS", "9h direct", 400, 4, "=D7*E7",
         "AC + SWISS daily direct", "https://www.aircanada.com"),
    ]

    for i, data in enumerate(flight_data):
        row = 4 + i
        route, airlines, duration, price, qty, formula, freq, source = data
        ws_flights.cell(row=row, column=1, value=route)
        ws_flights.cell(row=row, column=2, value=airlines)
        ws_flights.cell(row=row, column=3, value=duration)
        ws_flights.cell(row=row, column=4, value=price)
        ws_flights.cell(row=row, column=5, value=qty)
        ws_flights.cell(row=row, column=6, value=formula)
        ws_flights.cell(row=row, column=7, value=freq)

        # Hyperlink for source
        cell = ws_flights.cell(row=row, column=8, value=source)
        cell.hyperlink = source
        cell.font = link_font

        for col in range(1, 9):
            style_data_cell(ws_flights, row, col, is_currency=(col in [4, 6]))

    # Total row
    total_row = 8
    ws_flights.cell(row=total_row, column=1, value="TOTAL FLIGHTS")
    ws_flights.cell(row=total_row, column=1).font = total_font
    ws_flights.cell(row=total_row, column=6, value="=SUM(F4:F7)")
    ws_flights.cell(row=total_row, column=6).number_format = currency_format
    ws_flights.cell(row=total_row, column=6).font = total_font
    for c in range(1, 9):
        ws_flights.cell(row=total_row, column=c).fill = total_fill
        ws_flights.cell(row=total_row, column=c).border = thin_border

    for col_letter in ["A", "B", "C", "D", "E", "F", "G", "H"]:
        ws_flights.column_dimensions[col_letter].width = 20 if col_letter != "B" else 38

    # ========================================
    # Sheet 3: Accommodation
    # ========================================
    ws_accom = wb.create_sheet("Accommodation")
    ws_accom.sheet_properties.tabColor = "ED7D31"

    ws_accom["A1"] = "Accommodation Budget - 20 Nights, 3 Hotel Check-ins"
    ws_accom["A1"].font = Font(name="Calibri", bold=True, size=14, color="2F5496")
    ws_accom.merge_cells("A1:I1")

    accom_headers = ["City", "Hotel", "Room Type", "Nights", "Rate/Night", "Total", "Breakfast", "Key Features", "Booking Link"]
    for col, h in enumerate(accom_headers, 1):
        ws_accom.cell(row=3, column=col, value=h)
    style_header_row(ws_accom, 3, len(accom_headers))

    accom_data = [
        ("Lisbon", "Hotel da Baixa ****", "Family Room (Queen+King)", 5, 110, "=D4*E4",
         "Included", "2-min walk to Rossio Station, balconies, Nespresso", "https://www.hoteldabaixa.com"),
        ("Porto", "Moov Hotel Porto Centro ****", "Superior (2 double+sofa)", 2, 100, "=D5*E5",
         "Not included", "5-min to Sao Bento Station, modern design", "https://www.moovhotels.com"),
        ("Barcelona", "Aparthotel Arai 4* Superior", "Family apartment", 7, 120, "=D6*E6",
         "Not included", "Kitchenette, rooftop pool, Gothic Quarter", "https://www.hotelarai.com"),
        ("Zurich", "Hotel Glockenhof ****", "Family Suite (connecting)", 6, 160, "=D7*E7",
         "Included", "7-min to Zurich HB, Swiss chocolate daily", "https://www.glockenhof.ch"),
    ]

    for i, data in enumerate(accom_data):
        row = 4 + i
        city, hotel, room, nights, rate, formula, bfast, features, link = data
        ws_accom.cell(row=row, column=1, value=city)
        ws_accom.cell(row=row, column=2, value=hotel)
        ws_accom.cell(row=row, column=3, value=room)
        ws_accom.cell(row=row, column=4, value=nights)
        ws_accom.cell(row=row, column=5, value=rate)
        ws_accom.cell(row=row, column=6, value=formula)
        ws_accom.cell(row=row, column=7, value=bfast)
        ws_accom.cell(row=row, column=8, value=features)

        cell = ws_accom.cell(row=row, column=9, value=link)
        cell.hyperlink = link
        cell.font = link_font

        for col in range(1, 10):
            style_data_cell(ws_accom, row, col, is_currency=(col in [5, 6]))

    # Total row
    total_row = 8
    ws_accom.cell(row=total_row, column=1, value="TOTAL ACCOMMODATION")
    ws_accom.cell(row=total_row, column=1).font = total_font
    ws_accom.cell(row=total_row, column=4, value="=SUM(D4:D7)")
    ws_accom.cell(row=total_row, column=4).font = total_font
    ws_accom.cell(row=total_row, column=6, value="=SUM(F4:F7)")
    ws_accom.cell(row=total_row, column=6).number_format = currency_format
    ws_accom.cell(row=total_row, column=6).font = total_font
    for c in range(1, 10):
        ws_accom.cell(row=total_row, column=c).fill = total_fill
        ws_accom.cell(row=total_row, column=c).border = thin_border

    # Booking tips
    tips_row = 10
    ws_accom.cell(row=tips_row, column=1, value="Booking Tips:")
    ws_accom.cell(row=tips_row, column=1).font = Font(name="Calibri", bold=True, size=11)
    tips = [
        "Book 2-3 months ahead for September (shoulder season pricing)",
        "Request connecting rooms or family rooms when booking",
        "Check hotel websites directly - often same price as Booking.com but more flexibility",
        "All recommended hotels are near train stations for easy day trips",
    ]
    for i, tip in enumerate(tips):
        ws_accom.cell(row=tips_row + 1 + i, column=1, value=f"  {i+1}. {tip}")
        ws_accom.merge_cells(start_row=tips_row + 1 + i, start_column=1, end_row=tips_row + 1 + i, end_column=9)

    for col_letter in ["A", "B", "C", "D", "E", "F", "G", "H", "I"]:
        ws_accom.column_dimensions[col_letter].width = 18 if col_letter not in ["B", "H", "I"] else 32

    # ========================================
    # Sheet 4: Transportation
    # ========================================
    ws_transport = wb.create_sheet("Transportation")
    ws_transport.sheet_properties.tabColor = "70AD47"

    ws_transport["A1"] = "Transportation Budget - Trains, Metro, Local Transit"
    ws_transport["A1"].font = Font(name="Calibri", bold=True, size=14, color="2F5496")
    ws_transport.merge_cells("A1:H1")

    transport_headers = ["Country", "Item", "Unit Cost", "Qty/People", "Total", "Route/Details", "Duration", "Source"]
    for col, h in enumerate(transport_headers, 1):
        ws_transport.cell(row=3, column=col, value=h)
    style_header_row(ws_transport, 3, len(transport_headers))

    transport_data = [
        ("Switzerland", "Swiss Travel Pass (8-day consecutive)", 380, 4, "=C4*D4",
         "Unlimited trains/buses/boats + 50% mountain railways", "8 days", "https://www.swisstravelsystem.com"),
        ("Portugal", "Lisbon -> Porto Alfa Pendular", 30, 4, "=C5*D5",
         "Santa Apolonia -> Porto Campanha", "3h", "https://www.cp.pt"),
        ("Portugal", "Lisbon -> Sintra train (RT)", 5, 4, "=C6*D6",
         "Rossio Station -> Sintra", "40 min", "https://www.cp.pt"),
        ("Portugal", "Lisbon -> Cascais train (RT)", 5, 4, "=C7*D7",
         "Cais do Sodre -> Cascais", "40 min", "https://www.cp.pt"),
        ("Portugal", "Lisboa Viva card + loads", 10, 4, "=C8*D8",
         "Metro, buses, ferries in Lisbon area", "5 days", "https://www.metrolisboa.pt"),
        ("Spain", "Barcelona T-Casual (10 trips)", 12, 4, "=C9*D9",
         "Metro/buses/trams in Barcelona zone", "10 trips each", "https://www.tmb.cat"),
        ("Spain", "Barcelona -> Blanes (Costa Brava) RT", 10, 4, "=C10*D10",
         "Rodalies R1", "1h 20min", "https://rofrens.gencat.cat"),
        ("Spain", "Barcelona -> Sitges RT", 8, 4, "=C11*D11",
         "Rodalies R2", "35 min", "https://rodalies.gencat.cat"),
        ("Spain", "Barcelona -> Girona AVE RT", 24, 4, "=C12*D12",
         "AVE high-speed", "40 min", "https://www.renfe.com"),
        ("Spain", "Montserrat Tot Complert RT", 25, 4, "=C13*D13",
         "FGC train + rack railway + cable car", "1h 15min", "https://www.fgc.cat"),
        ("Spain", "Tibidabo transport RT", 12, 4, "=C14*D14",
         "Metro L7 + Bus 196 + Funicular", "45 min", "https://www.tmb.cat"),
    ]

    for i, data in enumerate(transport_data):
        row = 4 + i
        country, item, unit_cost, qty, formula, details, duration, source = data
        ws_transport.cell(row=row, column=1, value=country)
        ws_transport.cell(row=row, column=2, value=item)
        ws_transport.cell(row=row, column=3, value=unit_cost)
        ws_transport.cell(row=row, column=4, value=qty)
        ws_transport.cell(row=row, column=5, value=formula)
        ws_transport.cell(row=row, column=6, value=details)
        ws_transport.cell(row=row, column=7, value=duration)

        cell = ws_transport.cell(row=row, column=8, value=source)
        cell.hyperlink = source
        cell.font = link_font

        for col in range(1, 9):
            style_data_cell(ws_transport, row, col, is_currency=(col in [3, 5]))

    total_row = 4 + len(transport_data)
    ws_transport.cell(row=total_row, column=1, value="TOTAL TRANSPORTATION")
    ws_transport.cell(row=total_row, column=1).font = total_font
    ws_transport.cell(row=total_row, column=5, value=f"=SUM(E4:E{total_row-1})")
    ws_transport.cell(row=total_row, column=5).number_format = currency_format
    ws_transport.cell(row=total_row, column=5).font = total_font
    for c in range(1, 9):
        ws_transport.cell(row=total_row, column=c).fill = total_fill
        ws_transport.cell(row=total_row, column=c).border = thin_border

    for col_letter in ["A", "B", "C", "D", "E", "F", "G", "H"]:
        ws_transport.column_dimensions[col_letter].width = 18 if col_letter not in ["B", "F"] else 40

    # ========================================
    # Sheet 5: Activities
    # ========================================
    ws_activities = wb.create_sheet("Activities")
    ws_activities.sheet_properties.tabColor = "FFC000"

    ws_activities["A1"] = "Activities Budget - Nature-Focused"
    ws_activities["A1"].font = Font(name="Calibri", bold=True, size=14, color="2F5496")
    ws_activities.merge_cells("A1:H1")

    activity_headers = ["Day", "Country", "Activity", "Cost/Person", "Qty", "Total", "Type", "Source"]
    for col, h in enumerate(activity_headers, 1):
        ws_activities.cell(row=3, column=col, value=h)
    style_header_row(ws_activities, 3, len(activity_headers))

    activity_data = [
        ("Day 14", "Spain", "Tibidabo viewing platform elevator", 5, 4, "=D4*E4",
         "Mountain viewpoint", "https://en.wikipedia.org/wiki/Tibidabo"),
        ("Day 17", "Switzerland", "Trummelbach Falls entrance", 15, 4, "=D5*E5",
         "Waterfall", "https://www.truemmelbachfaelle.ch"),
        ("Day 18", "Switzerland", "Grindelwald First cable car (50% off w/ Swiss Pass)", 50, 4, "=D6*E6",
         "Mountain cable car", "https://www.jungfrau.ch"),
        ("Day 19", "Switzerland", "Mount Pilatus cable car (50% off w/ Swiss Pass, optional)", 25, 4, "=D7*E7",
         "Mountain cable car", "https://www.pilatus.ch"),
        ("Day 20", "Switzerland", "Rhine Falls boat ride to rock", 7, 4, "=D8*E8",
         "Boat ride", "https://www.rheinfall.ch"),
    ]

    for i, data in enumerate(activity_data):
        row = 4 + i
        day, country, activity, cost_pp, qty, formula, act_type, source = data
        ws_activities.cell(row=row, column=1, value=day)
        ws_activities.cell(row=row, column=2, value=country)
        ws_activities.cell(row=row, column=3, value=activity)
        ws_activities.cell(row=row, column=4, value=cost_pp)
        ws_activities.cell(row=row, column=5, value=qty)
        ws_activities.cell(row=row, column=6, value=formula)
        ws_activities.cell(row=row, column=7, value=act_type)

        cell = ws_activities.cell(row=row, column=8, value=source)
        cell.hyperlink = source
        cell.font = link_font

        for col in range(1, 9):
            style_data_cell(ws_activities, row, col, is_currency=(col in [4, 6]))

    total_row = 4 + len(activity_data)
    ws_activities.cell(row=total_row, column=1, value="TOTAL ACTIVITIES")
    ws_activities.cell(row=total_row, column=1).font = total_font
    ws_activities.cell(row=total_row, column=6, value=f"=SUM(F4:F{total_row-1})")
    ws_activities.cell(row=total_row, column=6).number_format = currency_format
    ws_activities.cell(row=total_row, column=6).font = total_font
    for c in range(1, 9):
        ws_activities.cell(row=total_row, column=c).fill = total_fill
        ws_activities.cell(row=total_row, column=c).border = thin_border

    for col_letter in ["A", "B", "C", "D", "E", "F", "G", "H"]:
        ws_activities.column_dimensions[col_letter].width = 18 if col_letter != "C" else 48

    # ========================================
    # Sheet 6: Food
    # ========================================
    ws_food = wb.create_sheet("Food")
    ws_food.sheet_properties.tabColor = "FF6B6B"

    ws_food["A1"] = "Food Budget - Estimated Daily Costs by City"
    ws_food["A1"].font = Font(name="Calibri", bold=True, size=14, color="2F5496")
    ws_food.merge_cells("A1:H1")

    ws_food["A2"] = "Note: Food costs come from the $1,410 buffer. Grocery stores recommended to stay within budget."
    ws_food["A2"].font = Font(name="Calibri", size=10, italic=True)
    ws_food.merge_cells("A2:H2")

    food_headers = ["City", "Days", "Daily Budget (4 ppl)", "Total", "Breakfast", "Grocery Stores", "Avg Meal Out", "Notes"]
    for col, h in enumerate(food_headers, 1):
        ws_food.cell(row=4, column=col, value=h)
    style_header_row(ws_food, 4, len(food_headers))

    food_data = [
        ("Lisbon", 5, 45, "=B5*C5", "Included at hotel",
         "Pingo Doce, Continente, Minipreco", "$8-12/meal",
         "Portugal is very affordable for food"),
        ("Porto", 2, 45, "=B6*C6", "At cafe or packed",
         "Pingo Doce, Continente", "$8-12/meal",
         "Seafood is excellent and affordable"),
        ("Barcelona", 7, 55, "=B7*C7", "Kitchenette available",
         "Mercadona, Bon Preu, Carrefour Express", "$10-15/meal",
         "Use kitchenette at Aparthotel Arai to save"),
        ("Zurich", 6, 75, "=B8*C8", "Included at hotel",
         "Migros (best value), Coop, Denner, Aldi", "$15-25/meal",
         "Switzerland is expensive - grocery shop at Migros"),
    ]

    for i, data in enumerate(food_data):
        row = 5 + i
        city, days, daily, formula, bfast, stores, avg_meal, notes = data
        ws_food.cell(row=row, column=1, value=city)
        ws_food.cell(row=row, column=2, value=days)
        ws_food.cell(row=row, column=3, value=daily)
        ws_food.cell(row=row, column=4, value=formula)
        ws_food.cell(row=row, column=5, value=bfast)
        ws_food.cell(row=row, column=6, value=stores)
        ws_food.cell(row=row, column=7, value=avg_meal)
        ws_food.cell(row=row, column=8, value=notes)

        for col in range(1, 9):
            style_data_cell(ws_food, row, col, is_currency=(col in [3, 4]))

    total_row = 9
    ws_food.cell(row=total_row, column=1, value="TOTAL FOOD ESTIMATE")
    ws_food.cell(row=total_row, column=1).font = total_font
    ws_food.cell(row=total_row, column=2, value="=SUM(B5:B8)")
    ws_food.cell(row=total_row, column=4, value="=SUM(D5:D8)")
    ws_food.cell(row=total_row, column=4).number_format = currency_format
    ws_food.cell(row=total_row, column=4).font = total_font
    for c in range(1, 9):
        ws_food.cell(row=total_row, column=c).fill = total_fill
        ws_food.cell(row=total_row, column=c).border = thin_border

    # Remaining after food
    row = total_row + 2
    ws_food.cell(row=row, column=1, value="Buffer Available")
    ws_food.cell(row=row, column=1).font = Font(name="Calibri", bold=True, size=11)
    ws_food.cell(row=row, column=4, value=1410)
    ws_food.cell(row=row, column=4).number_format = currency_format

    row += 1
    ws_food.cell(row=row, column=1, value="Remaining After Food")
    ws_food.cell(row=row, column=1).font = Font(name="Calibri", bold=True, size=11, color="006100")
    ws_food.cell(row=row, column=4, value=f"=D{total_row+2}-D{total_row}")
    ws_food.cell(row=row, column=4).number_format = currency_format
    ws_food.cell(row=row, column=4).font = Font(name="Calibri", bold=True, size=11, color="006100")

    for col_letter in ["A", "B", "C", "D", "E", "F", "G", "H"]:
        ws_food.column_dimensions[col_letter].width = 18 if col_letter not in ["F", "H"] else 38

    # Save
    output_path = "/home/user/ii-agent/travel/europe-2025/budget/budget_v1.xlsx"
    wb.save(output_path)
    print(f"Budget saved to: {output_path}")
    return output_path


if __name__ == "__main__":
    create_budget()
