"""Generate a comprehensive Europe Trip 2025 PDF summary using reportlab."""

from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.colors import HexColor, white, black
from reportlab.lib.units import inch
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak,
    KeepTogether,
)

OUTPUT = "/home/user/ii-agent/travel/europe-2025/Europe_Trip_2025_Summary.pdf"

# Colors
BLUE = HexColor("#2F5496")
LIGHT_BLUE = HexColor("#D6E4F0")
GREEN = HexColor("#006100")
LIGHT_GREEN = HexColor("#C6EFCE")
DARK = HexColor("#282828")
GRAY = HexColor("#666666")
RED_DARK = HexColor("#B41414")
ORANGE = HexColor("#C85000")
CREAM = HexColor("#FFF8DC")
WHITE = white

styles = getSampleStyleSheet()

# Custom styles
styles.add(ParagraphStyle("CoverTitle", parent=styles["Title"], fontSize=28, textColor=BLUE, spaceAfter=6, alignment=TA_CENTER))
styles.add(ParagraphStyle("CoverSub", parent=styles["Normal"], fontSize=15, textColor=GRAY, alignment=TA_CENTER, spaceAfter=10))
styles.add(ParagraphStyle("SectionTitle", parent=styles["Heading1"], fontSize=15, textColor=BLUE, spaceBefore=16, spaceAfter=6, borderWidth=1, borderColor=BLUE, borderPadding=4))
styles.add(ParagraphStyle("SubTitle", parent=styles["Heading2"], fontSize=12, textColor=HexColor("#404040"), spaceBefore=10, spaceAfter=4))
styles.add(ParagraphStyle("Body", parent=styles["Normal"], fontSize=10, textColor=DARK, leading=14, spaceAfter=4))
styles.add(ParagraphStyle("TripBullet", parent=styles["Normal"], fontSize=10, textColor=DARK, leading=14, leftIndent=20, bulletIndent=8, spaceAfter=2))
styles.add(ParagraphStyle("SmallItalic", parent=styles["Normal"], fontSize=9, textColor=GRAY, leading=12, alignment=TA_CENTER, spaceAfter=6))
styles.add(ParagraphStyle("DayHeader", parent=styles["Heading3"], fontSize=11, textColor=WHITE, spaceBefore=8, spaceAfter=4))
styles.add(ParagraphStyle("FactLabel", parent=styles["Normal"], fontSize=11, textColor=BLUE, leading=16))
styles.add(ParagraphStyle("FactValue", parent=styles["Normal"], fontSize=11, textColor=DARK, leading=16))
styles.add(ParagraphStyle("TimeBlock", parent=styles["Normal"], fontSize=9, textColor=DARK, leading=12, leftIndent=10))


def make_table(headers, rows, col_widths=None, header_color=BLUE):
    data = [headers] + rows
    t = Table(data, colWidths=col_widths, repeatRows=1)
    style_cmds = [
        ("BACKGROUND", (0, 0), (-1, 0), header_color),
        ("TEXTCOLOR", (0, 0), (-1, 0), WHITE),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, 0), 9),
        ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 1), (-1, -1), 9),
        ("ALIGN", (0, 0), (-1, 0), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.5, HexColor("#CCCCCC")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [WHITE, LIGHT_BLUE]),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
    ]
    t.setStyle(TableStyle(style_cmds))
    return t


def make_total_row(values, col_widths=None):
    data = [values]
    t = Table(data, colWidths=col_widths)
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), LIGHT_GREEN),
        ("TEXTCOLOR", (0, 0), (-1, 0), GREEN),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, 0), 9),
        ("GRID", (0, 0), (-1, -1), 0.5, HexColor("#CCCCCC")),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
    ]))
    return t


def make_day_block(day_num, title, blocks, color=BLUE):
    """Create a day block with time entries."""
    elements = []
    # Day header as a colored table cell
    day_hdr = Table([[f"  Day {day_num} - {title}"]], colWidths=[7.3 * inch])
    day_hdr.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (0, 0), color),
        ("TEXTCOLOR", (0, 0), (0, 0), WHITE),
        ("FONTNAME", (0, 0), (0, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (0, 0), 10),
        ("TOPPADDING", (0, 0), (0, 0), 4),
        ("BOTTOMPADDING", (0, 0), (0, 0), 4),
    ]))
    elements.append(day_hdr)

    # Time block table
    data = [["Time", "Activity", "Meal", "Transport", "Cost"]]
    for block_name, activity, meal, transport, cost in blocks:
        data.append([block_name, activity, meal, transport, f"${cost}"])

    t = Table(data, colWidths=[1.0*inch, 2.3*inch, 1.5*inch, 1.2*inch, 0.6*inch], repeatRows=1)
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), HexColor("#E8E8E8")),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
        ("GRID", (0, 0), (-1, -1), 0.3, HexColor("#DDDDDD")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("ALIGN", (-1, 0), (-1, -1), "RIGHT"),
    ]))
    elements.append(t)
    elements.append(Spacer(1, 6))
    return KeepTogether(elements)


def build_pdf():
    doc = SimpleDocTemplate(
        OUTPUT,
        pagesize=letter,
        topMargin=0.6*inch,
        bottomMargin=0.6*inch,
        leftMargin=0.6*inch,
        rightMargin=0.6*inch,
    )
    story = []

    # ========================= COVER PAGE =========================
    story.append(Spacer(1, 1.5*inch))
    story.append(Paragraph("Europe Trip 2025", styles["CoverTitle"]))
    story.append(Paragraph("Family Travel Plan", styles["CoverSub"]))
    story.append(Spacer(1, 0.3*inch))

    # Quick facts table
    facts = [
        ["Travelers:", "4 people (family group)"],
        ["Duration:", "21 days / 3 weeks"],
        ["Dates:", "September 2025"],
        ["Budget:", "$10,000 USD total"],
        ["Origin:", "Toronto, Ontario, Canada"],
        ["Route:", "Toronto > Lisbon > Porto > Barcelona > Zurich > Toronto"],
        ["Hotel Check-ins:", "3 total (minimizing moves)"],
        ["Flights:", "All direct - no layovers"],
    ]
    ft = Table(facts, colWidths=[1.5*inch, 4.5*inch])
    ft.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 11),
        ("TEXTCOLOR", (0, 0), (0, -1), BLUE),
        ("TEXTCOLOR", (1, 0), (1, -1), DARK),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ("ALIGN", (0, 0), (0, -1), "RIGHT"),
    ]))
    story.append(ft)

    story.append(Spacer(1, 0.4*inch))
    story.append(Paragraph(
        '<i>"Slow travel focused on natural beauty, outdoor experiences, and family relaxation '
        '- avoiding tourist trap interiors and museum fatigue while maximizing coastal and alpine '
        'scenery with minimal logistical stress."</i>',
        styles["SmallItalic"]
    ))

    # ========================= OVERVIEW =========================
    story.append(PageBreak())
    story.append(Paragraph("Trip Overview", styles["SectionTitle"]))

    story.append(Paragraph("Travel Preferences", styles["SubTitle"]))
    for w in [
        "Nature-focused: beaches, mountains, hiking, coastal walks, waterfalls, viewpoints",
        "All direct flights only - no layovers, no connections",
        "Minimal hotel changes - base cities with day trips",
        "Maximum 1 major activity per day, relaxed pace",
        "Train travel for intercity and day trips",
        "Family-friendly accommodations with connecting rooms / family suites",
    ]:
        story.append(Paragraph(w, styles["TripBullet"], bulletText="\u2022"))

    story.append(Paragraph("What This Trip Does NOT Include", styles["SubTitle"]))
    for n in [
        "No museums - zero museum visits",
        "No building interiors - no palaces, churches, or castles (exteriors OK)",
        "No layover flights - direct only",
        "No constant hotel hopping - 3 check-ins for 21 days",
    ]:
        story.append(Paragraph(n, styles["TripBullet"], bulletText="\u2717"))

    story.append(Paragraph("Route Summary", styles["SubTitle"]))
    story.append(Paragraph(
        "<b>Week 1 - Portugal:</b> Lisbon (5 nights) + Porto (2 nights) - Atlantic coast, cliffs, beaches<br/>"
        "<b>Week 2 - Spain:</b> Barcelona (7 nights) - Mediterranean coast, mountains, beach towns<br/>"
        "<b>Week 3 - Switzerland:</b> Zurich (6 nights) - Alpine waterfalls, lakes, mountain hiking",
        styles["Body"]
    ))

    # ========================= FLIGHTS =========================
    story.append(Paragraph("Flights (All Direct)", styles["SectionTitle"]))
    cw = [1.2*inch, 1.8*inch, 0.8*inch, 0.9*inch, 0.9*inch]
    story.append(make_table(
        ["Route", "Airlines", "Duration", "Per Person", "Total (4)"],
        [
            ["Toronto > Lisbon", "TAP / Air Canada / Transat", "7h", "$450", "$1,800"],
            ["Lisbon > Barcelona", "TAP / Vueling / Ryanair", "2h", "$50", "$200"],
            ["Barcelona > Zurich", "SWISS / Vueling", "1h 45m", "$60", "$240"],
            ["Zurich > Toronto", "Air Canada / SWISS", "9h", "$400", "$1,600"],
        ],
        col_widths=cw
    ))
    story.append(make_total_row(["TOTAL", "", "", "", "$3,840"], col_widths=cw))

    story.append(Spacer(1, 6))
    story.append(Paragraph("Booking Notes", styles["SubTitle"]))
    for note in [
        "Book 2-3 months before September for best prices",
        "TAP Air Portugal: 8 nonstop flights/week Toronto-Lisbon",
        "Lisbon-Barcelona: 12+ direct flights daily (budget airlines available)",
        "Air Transat Toronto-Lisbon is seasonal - verify September availability",
    ]:
        story.append(Paragraph(note, styles["TripBullet"], bulletText="\u2022"))

    # ========================= ACCOMMODATION =========================
    story.append(Paragraph("Accommodation (20 Nights, 3 Check-ins)", styles["SectionTitle"]))
    cw2 = [0.7*inch, 1.5*inch, 1.2*inch, 0.6*inch, 0.7*inch, 0.6*inch, 1.7*inch]
    story.append(make_table(
        ["City", "Hotel", "Room Type", "Nights", "Rate", "Total", "Key Features"],
        [
            ["Lisbon", "Hotel da Baixa ****", "Family (Queen+King)", "5", "$110", "$550", "2-min to Rossio, breakfast"],
            ["Porto", "Moov Hotel ****", "Superior (2dbl+sofa)", "2", "$100", "$200", "5-min to Sao Bento"],
            ["Barcelona", "Aparthotel Arai 4*S", "Family apartment", "7", "$120", "$840", "Kitchenette, rooftop pool"],
            ["Zurich", "Hotel Glockenhof ****", "Family Suite", "6", "$160", "$960", "Breakfast, Swiss chocolate"],
        ],
        col_widths=cw2
    ))
    story.append(make_total_row(["", "TOTAL", "", "20 nts", "", "$2,550", ""], col_widths=cw2))

    # ========================= BUDGET =========================
    story.append(PageBreak())
    story.append(Paragraph("Budget Breakdown ($10,000 Total)", styles["SectionTitle"]))

    cw3 = [2.2*inch, 1.0*inch, 4.1*inch]
    story.append(make_table(
        ["Category", "Cost (4 ppl)", "Notes"],
        [
            ["Flights", "$3,840", "All direct: TAP, Air Canada, SWISS, Vueling"],
            ["Accommodation", "$2,550", "20 nights across 4 hotels (3 check-ins)"],
            ["Swiss Travel Pass (8-day x4)", "$1,520", "Unlimited Swiss trains/buses/boats + 50% mountain railways"],
            ["Portugal trains/metro", "$140", "Alfa Pendular Lisbon-Porto + local transit"],
            ["Spain trains/metro", "$140", "Barcelona day trips + T-Casual card"],
            ["Activities", "$400", "Cable cars, waterfall entries, boat rides"],
        ],
        col_widths=cw3
    ))
    story.append(make_total_row(["GRAND TOTAL", "$8,590", "85.9% of budget allocated"], col_widths=cw3))

    # Buffer row
    buf = Table([["REMAINING BUFFER", "$1,410", "For food (~$55/day avg), emergencies, extras"]], colWidths=cw3)
    buf.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), CREAM),
        ("TEXTCOLOR", (0, 0), (-1, 0), GREEN),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, 0), 9),
        ("GRID", (0, 0), (-1, -1), 0.5, HexColor("#CCCCCC")),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
    ]))
    story.append(buf)

    story.append(Spacer(1, 10))
    story.append(Paragraph("Food Budget Estimate (from buffer)", styles["SubTitle"]))
    cw4 = [1.0*inch, 0.6*inch, 1.0*inch, 0.7*inch, 3.0*inch]
    story.append(make_table(
        ["City", "Days", "Daily (4 ppl)", "Total", "Grocery Stores"],
        [
            ["Lisbon", "5", "$45", "$225", "Pingo Doce, Continente, Minipreco"],
            ["Porto", "2", "$45", "$90", "Pingo Doce, Continente"],
            ["Barcelona", "7", "$55", "$385", "Mercadona, Carrefour Express"],
            ["Zurich", "6", "$75", "$450", "Migros (best), Coop, Aldi"],
        ],
        col_widths=cw4
    ))
    story.append(make_total_row(["TOTAL", "20", "", "$1,150", "Remaining after food: ~$260"], col_widths=cw4))

    # ========================= WEEK 1: PORTUGAL =========================
    story.append(PageBreak())
    story.append(Paragraph("Week 1: Portugal (Days 1-7)", styles["SectionTitle"]))
    story.append(Paragraph("Base: Lisbon (5 nights) + Porto (2 nights) | Focus: Atlantic coast, cliffs, beaches", styles["Body"]))

    W1_COLOR = HexColor("#007838")
    days_w1 = [
        (1, "Arrive Lisbon", [
            ("Morning", "In-flight from Toronto (7h direct)", "Airline meal", "TAP/Air Canada", "0"),
            ("Afternoon", "Check into Hotel da Baixa, rest", "Light lunch in Baixa", "Metro from airport", "15"),
            ("Evening", "Tagus River walk, sunset Praca do Comercio", "Riverside dinner", "Walking", "20"),
        ]),
        (2, "Sintra Coast Day Trip", [
            ("Morning", "Cabo da Roca - westernmost point of Europe", "Hotel breakfast (incl.)", "Train 40min", "10"),
            ("Afternoon", "Hike to Praia da Ursa - wild beach, rocks", "Packed lunch/cafe", "Local bus", "20"),
            ("Evening", "Praia do Guincho - dunes, surfing beach", "Seafood dinner", "Train back", "10"),
        ]),
        (3, "Cascais & Portuguese Riviera", [
            ("Morning", "Boca do Inferno sea cliffs", "Hotel breakfast (incl.)", "Train 40min", "10"),
            ("Afternoon", "Praia da Duquesa beach, coastal cycling", "Seafront restaurant", "Walk/cycle", "15"),
            ("Evening", "Sunset on Cascais promenade", "Sunset seafood dinner", "Train back", "10"),
        ]),
        (4, "Arrabida Natural Park", [
            ("Morning", "Mediterranean forest hike", "Hotel breakfast (incl.)", "Bus/taxi/rental", "25"),
            ("Afternoon", "Praia de Galapinhos - crystal swimming", "Packed lunch", "Within park", "10"),
            ("Evening", "Scenic drive back, relax", "Dinner in Baixa", "Return transport", "15"),
        ]),
        (5, "Lisbon Miradouros & Tagus River", [
            ("Morning", "Viewpoints: Senhora do Monte, Graca", "Hotel breakfast (incl.)", "Walking", "0"),
            ("Afternoon", "Santa Luzia viewpoint, ferry to Cacilhas", "Cafe lunch", "Walk + ferry", "10"),
            ("Evening", "Relaxed evening, pack for Porto", "Farewell Lisbon dinner", "Walking", "5"),
        ]),
        (6, "Train to Porto + Douro River", [
            ("Morning", "Alfa Pendular train (3h scenic ride)", "Hotel breakfast (incl.)", "Train", "35"),
            ("Afternoon", "Check in Moov Hotel, Ribeira waterfront", "Lunch overlooking Douro", "Walking", "0"),
            ("Evening", "Sunset Dom Luis I Bridge, Foz do Douro", "Dinner in Foz do Douro", "Walking", "0"),
        ]),
        (7, "Porto Beaches & Nature", [
            ("Morning", "Praia de Matosinhos - urban beach", "Beachside cafe", "Metro/bus", "5"),
            ("Afternoon", "Walk Atlantic coast to Foz do Douro", "Seaside restaurant", "Coastal path", "10"),
            ("Evening", "Palacio de Cristal gardens", "Dinner historic center", "Walking", "10"),
        ]),
    ]
    for d in days_w1:
        story.append(make_day_block(*d, color=W1_COLOR))

    # ========================= WEEK 2: SPAIN =========================
    story.append(PageBreak())
    story.append(Paragraph("Week 2: Spain (Days 8-15)", styles["SectionTitle"]))
    story.append(Paragraph("Base: Barcelona (7 nights) | Focus: Mediterranean coast, mountains, beach towns", styles["Body"]))

    W2_COLOR = HexColor("#C85000")
    days_w2 = [
        (8, "Fly Porto to Barcelona", [
            ("Morning", "Flight Porto-Barcelona (direct, 2h)", "Hotel breakfast", "Flight", "60"),
            ("Afternoon", "Check into Aparthotel Arai, Gothic Qtr", "Tapas in El Born", "Aerobus", "15"),
            ("Evening", "Stroll La Barceloneta beach", "Tapas dinner El Born", "Walking", "0"),
        ]),
        (9, "Barceloneta & Waterfront", [
            ("Morning", "Barceloneta Beach swimming", "Kitchenette breakfast", "Walking", "5"),
            ("Afternoon", "Port Olimpic + Parc Ciutadella", "Chiringuito lunch", "Walking", "10"),
            ("Evening", "Sunset W Hotel beach area", "Waterfront dinner", "Walking", "10"),
        ]),
        (10, "Costa Brava: Tossa de Mar", [
            ("Morning", "Train+bus to Tossa de Mar", "Hotel breakfast", "R1 train+bus", "10"),
            ("Afternoon", "Platja Gran, ancient walls, cove hike", "Old town lunch", "Walking", "10"),
            ("Evening", "Explore & return to Barcelona", "Dinner", "Bus+train", "10"),
        ]),
        (11, "Montserrat Mountain", [
            ("Morning", "FGC+rack railway to Montserrat", "Hotel breakfast", "FGC+rack", "25"),
            ("Afternoon", "Sant Joan trail - panoramic views", "Mountain restaurant", "Hiking", "10"),
            ("Evening", "Return to Barcelona", "Dinner in Barcelona", "Train back", "0"),
        ]),
        (12, "Sitges Beach Town", [
            ("Morning", "Train to Sitges beach town", "Hotel breakfast", "R2 train 35min", "5"),
            ("Afternoon", "Platja de Sant Sebastia, promenade", "Beachfront restaurant", "Walking", "10"),
            ("Evening", "Sunset beach drinks, return", "Dinner", "Train back", "10"),
        ]),
        (13, "Girona & Costa Brava", [
            ("Morning", "AVE to Girona, Onyar River walk", "Hotel breakfast", "AVE 40min", "15"),
            ("Afternoon", "Calella de Palafrugell fishing village", "Seaside restaurant", "Bus", "15"),
            ("Evening", "Return to Barcelona", "Dinner in Barcelona", "AVE train", "10"),
        ]),
        (14, "Mount Tibidabo", [
            ("Morning", "Metro+bus+funicular to summit (512m)", "Hotel breakfast", "Transit", "15"),
            ("Afternoon", "Sky Walk panoramic, Collserola option", "Mountain cafe", "Walking", "5"),
            ("Evening", "Return, rooftop pool at hotel", "Gothic Quarter dinner", "Transit", "5"),
        ]),
        (15, "Beach Day + Travel to Switzerland", [
            ("Morning", "Final beach at Barceloneta", "Kitchenette breakfast", "Walking", "0"),
            ("Afternoon", "Flight Barcelona-Zurich (1h45m)", "Airport lunch", "Flight", "60"),
            ("Evening", "Check in Glockenhof, Lake Zurich walk", "Old Town dinner", "Train 10min", "0"),
        ]),
    ]
    for d in days_w2:
        story.append(make_day_block(*d, color=W2_COLOR))

    # ========================= WEEK 3: SWITZERLAND =========================
    story.append(PageBreak())
    story.append(Paragraph("Week 3: Switzerland (Days 16-21)", styles["SectionTitle"]))
    story.append(Paragraph("Base: Zurich (6 nights) | Focus: Alpine waterfalls, lakes, mountain hiking", styles["Body"]))

    # Swiss Pass callout
    sp = Table([["SWISS TRAVEL PASS (8-Day) - $380/person x 4 = $1,520 | All trains/buses/boats FREE + 50% off mountain railways"]], colWidths=[7.3*inch])
    sp.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (0, 0), CREAM),
        ("TEXTCOLOR", (0, 0), (0, 0), HexColor("#8B6914")),
        ("FONTNAME", (0, 0), (0, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (0, 0), 9),
        ("TOPPADDING", (0, 0), (0, 0), 5),
        ("BOTTOMPADDING", (0, 0), (0, 0), 5),
        ("LEFTPADDING", (0, 0), (0, 0), 8),
    ]))
    story.append(sp)
    story.append(Spacer(1, 6))

    W3_COLOR = HexColor("#B41414")
    days_w3 = [
        (16, "Zurich Lake & Old Town", [
            ("Morning", "Old Town walk, Lindenhof viewpoint", "Hotel buffet (incl.)", "Walking", "0"),
            ("Afternoon", "Swim/relax Zurichhorn Park", "Lakeside cafe", "Walk/tram", "10"),
            ("Evening", "Sunset boat cruise (FREE Swiss Pass)", "Old Town dinner", "Boat+walk", "10"),
        ]),
        (17, "Lauterbrunnen Valley", [
            ("Morning", "Train to Lauterbrunnen - 72 Waterfalls", "Hotel buffet (incl.)", "Rail FREE", "0"),
            ("Afternoon", "Staubbach Falls + Trummelbach Falls", "Village lunch", "Walking", "60"),
            ("Evening", "Alpine meadow walk, train back", "Dinner in Zurich", "Rail FREE", "20"),
        ]),
        (18, "Grindelwald & First (2,168m)", [
            ("Morning", "Train+cable car to First (50% off)", "Hotel buffet (incl.)", "Rail FREE+cable", "100"),
            ("Afternoon", "Hike Lake Bachalpsee, First Cliff Walk", "Packed/mountain lunch", "Hiking", "0"),
            ("Evening", "Cable car down, train back", "Dinner in Zurich", "Rail FREE", "0"),
        ]),
        (19, "Lucerne & Lake", [
            ("Morning", "Train to Lucerne, Chapel Bridge", "Hotel buffet (incl.)", "Rail FREE", "0"),
            ("Afternoon", "Boat cruise FREE, opt. Mt. Pilatus 50%", "Lunch Lucerne/boat", "Boat FREE", "50"),
            ("Evening", "Lakefront walk, train back", "Dinner Lucerne/Zurich", "Rail FREE", "0"),
        ]),
        (20, "Rhine Falls", [
            ("Morning", "Train to Schaffhausen, Rhine Falls walk", "Hotel buffet (incl.)", "Rail FREE", "0"),
            ("Afternoon", "Boat to rock in falls, riverside walk", "Rhine Falls restaurant", "Boat", "25"),
            ("Evening", "Return to Zurich, farewell dinner", "Special farewell dinner", "Rail FREE", "0"),
        ]),
        (21, "Departure Day", [
            ("Morning", "Leisure morning, final lake walk, pack", "Hotel buffet (incl.)", "Walking", "0"),
            ("Afternoon", "Train to airport FREE, flight to Toronto 9h", "Airport lunch", "Rail FREE", "0"),
            ("Evening", "In-flight / Arrive Toronto", "Airline meal", "Flight", "0"),
        ]),
    ]
    for d in days_w3:
        story.append(make_day_block(*d, color=W3_COLOR))

    # ========================= NATURE HIGHLIGHTS =========================
    story.append(PageBreak())
    story.append(Paragraph("Nature Highlights by Country", styles["SectionTitle"]))

    story.append(Paragraph("Portugal - Coastal Focus", styles["SubTitle"]))
    for h in [
        "Cabo da Roca - Dramatic cliffs, westernmost point of continental Europe",
        "Praia da Ursa - Wild beach with massive rock formations",
        "Praia do Guincho - Windswept dunes, popular surfing beach",
        "Arrabida Natural Park - Mediterranean forest, pristine beaches, crystal water",
        "Douro River in Porto - UNESCO waterfront, bridge views",
        "Atlantic Coastline - Matosinhos beach, Foz do Douro district",
    ]:
        story.append(Paragraph(h, styles["TripBullet"], bulletText="\u2022"))

    story.append(Paragraph("Spain - Coastal + Mountain", styles["SubTitle"]))
    for h in [
        "Costa Brava - Hidden coves, crystal water, fishing villages",
        "Montserrat - Dramatic rock formations, hiking trails, panoramic views",
        "Mount Tibidabo - 512m summit, 360-degree Barcelona/sea/Pyrenees views",
        "Sitges & Barceloneta - Classic Mediterranean beaches",
        "Girona's Onyar River - Colorful riverside houses",
    ]:
        story.append(Paragraph(h, styles["TripBullet"], bulletText="\u2022"))

    story.append(Paragraph("Switzerland - Alpine Focus", styles["SubTitle"]))
    for h in [
        "Lauterbrunnen Valley - 72 waterfalls, alpine meadows, Staubbach Falls",
        "Trummelbach Falls - Glacial waterfalls inside a mountain",
        "Grindelwald First - Eiger/Monch/Jungfrau views, Lake Bachalpsee, Cliff Walk",
        "Lake Lucerne - Boat cruises with mountain backdrop",
        "Rhine Falls - Europe's largest waterfall, boat to central rock",
        "Lake Zurich - Swimming, boat cruises, lakeside parks",
    ]:
        story.append(Paragraph(h, styles["TripBullet"], bulletText="\u2022"))

    # ========================= TRAINS =========================
    story.append(Paragraph("Key Train Information", styles["SectionTitle"]))

    story.append(Paragraph("Portugal", styles["SubTitle"]))
    for t in [
        "Lisbon > Sintra: 40 min from Rossio Station (\u20ac2.30)",
        "Lisbon > Cascais: 40 min from Cais do Sodre (\u20ac2.30)",
        "Lisbon > Porto: 3h Alfa Pendular from Santa Apolonia (\u20ac27-34)",
        "Tip: Buy Lisboa Viva card for local transit",
    ]:
        story.append(Paragraph(t, styles["TripBullet"], bulletText="\u2022"))

    story.append(Paragraph("Spain", styles["SubTitle"]))
    for t in [
        "Barcelona > Blanes (Costa Brava): 1h 20m, Rodalies R1 (~\u20ac5)",
        "Barcelona > Sitges: 35 min, Rodalies R2 (~\u20ac4)",
        "Barcelona > Girona: 40 min, AVE high-speed (~\u20ac12)",
        "Barcelona > Montserrat: 1h 15m, FGC + rack railway (~\u20ac25 RT)",
        "Tip: Buy T-Casual card (10 trips) for Barcelona metro/trains",
    ]:
        story.append(Paragraph(t, styles["TripBullet"], bulletText="\u2022"))

    story.append(Paragraph("Switzerland (ALL covered by Swiss Travel Pass)", styles["SubTitle"]))
    for t in [
        "Zurich > Lauterbrunnen: 2.5h (change Interlaken) - FREE",
        "Zurich > Grindelwald: 2.5h (change Interlaken) - FREE",
        "Zurich > Lucerne: 45 min direct - FREE",
        "Zurich > Schaffhausen (Rhine Falls): 40 min direct - FREE",
        "Zurich HB > Zurich Airport: 10 min direct - FREE",
    ]:
        story.append(Paragraph(t, styles["TripBullet"], bulletText="\u2022"))

    # ========================= STILL TO DO =========================
    story.append(PageBreak())
    story.append(Paragraph("What Still Needs Planning / Booking", styles["SectionTitle"]))
    for item in [
        "Book flights (2-3 months before September 2025)",
        "Book hotels (2-3 months ahead for September shoulder season)",
        "Purchase Swiss Travel Pass (swisstravelsystem.com or Zurich HB station)",
        "Reserve specific restaurants if desired",
        "Book rental car/taxi for Arrabida Natural Park (Day 4)",
        "Confirm train schedules closer to travel dates",
        "Create packing list (beach gear + hiking gear + layers for Switzerland)",
        "Set up Google Maps saved locations for each destination",
        "Research travel insurance options",
        "Create rain-day backup plans for outdoor activities",
    ]:
        story.append(Paragraph(item, styles["TripBullet"], bulletText="\u2610"))

    story.append(Spacer(1, 12))
    story.append(Paragraph("Grocery Stores by City", styles["SubTitle"]))
    cw5 = [1.2*inch, 5.0*inch]
    story.append(make_table(
        ["City", "Stores (Best Value First)"],
        [
            ["Lisbon", "Pingo Doce (most popular), Continente, Minipreco (budget)"],
            ["Porto", "Pingo Doce, Continente"],
            ["Barcelona", "Mercadona (best value), Bon Preu, Carrefour Express"],
            ["Zurich", "Migros (BEST VALUE), Coop, Denner (discount), Aldi (budget)"],
        ],
        col_widths=cw5
    ))

    story.append(Spacer(1, 12))
    story.append(Paragraph(
        "Tip: Use hotel kitchenettes (especially Barcelona's Aparthotel Arai) to prepare "
        "breakfasts and snacks. This is especially important in Switzerland where eating "
        "out is 2-3x more expensive than Portugal or Spain.",
        styles["Body"]
    ))

    doc.build(story)
    print(f"PDF saved to: {OUTPUT}")


if __name__ == "__main__":
    build_pdf()
