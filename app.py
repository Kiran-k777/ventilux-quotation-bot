import streamlit as st
import pandas as pd
import re
from io import BytesIO
from openpyxl import load_workbook
from difflib import SequenceMatcher
from datetime import date
from PIL import Image
from PIL import ImageDraw
from PIL import ImageFont
import pymupdf as fitz


from pathlib import Path

# Streamlit Community Cloud: use the directory containing this app.py
# for bundled Excel and PDF template files.
BASE_DIR = Path(__file__).resolve().parent

def get_file(filename):
    """Find a bundled application file in Streamlit Cloud or local execution."""
    possible_paths = [
        BASE_DIR / filename,
        Path.cwd() / filename,
    ]

    for path in possible_paths:
        if path.exists():
            return path

    raise FileNotFoundError(
        f"Could not find '{filename}'. Checked: "
        f"{[str(p) for p in possible_paths]}"
    )

from reportlab.platypus import (
    SimpleDocTemplate,
    Table,
    TableStyle,
    Paragraph,
    Spacer
)

from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.pagesizes import A4


EURO_TO_GBP = 0.84


st.set_page_config(
    page_title="Quotation Bot",
    page_icon="💡",
    layout="wide"
)



# ==========================
# LOAD DATA
# ==========================

@st.cache_data
def load_data(customer_type):

    if customer_type == "Contractor":

        file=get_file("Base Final.xlsx")

    elif customer_type == "Wholesaler":

        file=get_file("Wholesaler 1.xlsx")

    elif customer_type == "OPW":

        file=get_file("OPW STD.xlsx")

    else:

        raise ValueError(
            f"Unknown customer type: {customer_type}"
        )


    df=pd.read_excel(file)

    df=df.dropna(how="all")

    df.columns=df.columns.str.strip()



    if customer_type in ["Contractor", "OPW"]:

        df.rename(
            columns={
                "Standalone":"Price_4_Years",
                "Standalone.1":"Price_8_Years"
            },
            inplace=True
        )


    else:

        standalone=[
            c for c in df.columns
            if str(c).startswith("Standalone")
        ]


        if len(standalone)>=4:

            df.rename(
                columns={

                    standalone[0]:"Price_4_1_49",
                    standalone[1]:"Price_4_50",
                    standalone[2]:"Price_8_1_49",
                    standalone[3]:"Price_8_50"

                },
                inplace=True
            )



    def clean_price(x):

        if isinstance(x,str):

            x=x.replace("€","")
            x=x.replace(",","")
            x=x.strip()

            if x in ["","NA","nan","N/A"]:

                return 0

        return x



    for c in df.columns:

        if "Price" in c:

            df[c]=df[c].apply(clean_price)

            df[c]=pd.to_numeric(
                df[c],
                errors="coerce"
            ).fillna(0)



    required=[

        "Price_4_Years",
        "Price_8_Years",
        "Price_4_1_49",
        "Price_4_50",
        "Price_8_1_49",
        "Price_8_50"

    ]


    for c in required:

        if c not in df.columns:

            df[c]=0



    search_cols=[

        "Fitting",
        "Product Code",
        "Type",
        "Mode",
        "Mounting",
        "Colour",
        "IP",
        "Lumens"

    ]


    for c in search_cols:

        if c not in df.columns:

            df[c]=""



    df["Product_Description"] = (
        df[search_cols]
        .apply(
            lambda row: " ".join(
                str(value).strip()
                for value in row
                if pd.notna(value) and str(value).strip()
            ),
            axis=1
        )
    )

    return df





# ==========================
# CLEAN CODE
# ==========================

def clean_code(x):

    return (

        str(x)
        .lower()
        .strip()
        .replace(" ","")
        .replace("/","")
        .replace("-","")
        .replace("_","")
        .replace(".","")
        .replace("\\","")

    )





# ==========================
# MATCH PRODUCT
# ==========================

def normalise_product_word(word):
    """
    Normalise common English plural forms so employee input such as
    'Fusions' matches the singular product name 'Fusion'.
    """
    word = str(word).lower().strip()

    if len(word) > 4 and word.endswith("ies"):
        return word[:-3] + "y"

    if len(word) > 4 and word.endswith("es"):
        stem = word[:-2]
        if stem.endswith(("s", "x", "z", "ch", "sh")):
            return stem

    if len(word) > 3 and word.endswith("s") and not word.endswith(("ss", "us", "is")):
        return word[:-1]

    return word


def normalise_product_text(value):
    """Lowercase text, remove punctuation and normalise plural words."""
    words = re.findall(r"[a-z0-9]+", str(value).lower())
    return " ".join(normalise_product_word(word) for word in words)


def match_product(text,df):

    search=text.lower()

    search=re.sub(
        r'^\s*\d+\s*x\s*',
        '',
        search
    ).strip()

    code_search=clean_code(search)

    df["clean_code"]=(
        df["Product Code"]
        .apply(clean_code)
    )

    # Keep the existing exact/partial product-code matching.
    result=df[
        df["clean_code"].str.contains(
            code_search,
            na=False
        )
    ]

    if len(result):
        return result.iloc[0]

    # Keep the existing fuzzy product-code matching.
    best=None
    best_score=0

    for _,row in df.iterrows():

        score=SequenceMatcher(
            None,
            code_search,
            row["clean_code"]
        ).ratio()

        if score>best_score:
            best_score=score
            best=row

    if best_score>=0.75:
        return best

    # Match descriptions after normalising singular/plural words.
    # Examples: "Fusions" -> "fusion", "Evoluxes" -> "evolux".
    search_words=normalise_product_text(search).split()
    search_set=set(search_words)

    temp=df.copy()
    temp["_normalised_description"]=(
        temp["Product_Description"].apply(normalise_product_text)
    )

    def description_score(description):
        description_words=set(description.split())
        return sum(word in description_words for word in search_set)

    temp["score"]=temp["_normalised_description"].apply(description_score)
    temp=temp.sort_values("score", ascending=False)

    if temp.empty or temp.iloc[0]["score"]==0:
        return None

    return temp.iloc[0]




# ==========================
# CREATE QUOTE
# ==========================

def create_quote(
    items,
    customer_type,
    warranty,
    refs,
    region
):


    df=load_data(customer_type)

    quote=[]



    for item in items:


        qty_match=re.search(

            r'(\d+)\s*x',
            item,
            re.I

        )


        qty=int(

            qty_match.group(1)

        ) if qty_match else 1



        product=match_product(
            item,
            df
        )



        if product is None:

            continue



        if customer_type in ["Contractor", "OPW"]:


            if warranty=="4 Years":

                price=product["Price_4_Years"]
                used="4 Years"


                if price==0:

                    price=product["Price_8_Years"]
                    used="8 Years"



            else:

                price=product["Price_8_Years"]
                used="8 Years"


                if price==0:

                    price=product["Price_4_Years"]
                    used="4 Years"




        else:

            # Wholesaler pricing has separate prices for:
            #   4 Years: 1-49 / 50+
            #   8 Years: 1-49 / 50+
            #
            # If the requested warranty is unavailable for the
            # selected quantity tier, automatically use the other
            # warranty for THE SAME quantity tier.
            #
            # Example:
            #   Customer asks for 4 Years, qty >= 50
            #   4-year 50+ price missing
            #   -> use 8-year 50+ price.
            #
            # This is important because falling back to the 1-49
            # price would give the wrong wholesaler price.

            if qty < 50:

                if warranty == "4 Years":

                    price = product["Price_4_1_49"]
                    used = "4 Years"

                    if price == 0:

                        price = product["Price_8_1_49"]
                        used = "8 Years"

                else:

                    price = product["Price_8_1_49"]
                    used = "8 Years"

                    if price == 0:

                        price = product["Price_4_1_49"]
                        used = "4 Years"

            else:

                if warranty == "4 Years":

                    price = product["Price_4_50"]
                    used = "4 Years"

                    if price == 0:

                        price = product["Price_8_50"]
                        used = "8 Years"

                else:

                    price = product["Price_8_50"]
                    used = "8 Years"

                    if price == 0:

                        price = product["Price_4_50"]
                        used = "4 Years"




        if region=="Northern Ireland":

            price=round(
                price*EURO_TO_GBP,
                2
            )

            currency="£"


        else:

            currency="€"



        quote.append({

            "Ref":
            refs.get(
                item,
                len(quote)+1
            ),

            "Product":
            product["Fitting"],

            "Product Code":
            product["Product Code"],

            "Type":
            product["Type"],

            "Mode":
            product["Mode"],

            "Mounting":
            product["Mounting"],

            "Colour":
            product["Colour"],

            "IP":
            product["IP"],

            "Lumens":
            product["Lumens"],

            "Warranty":
            used,

            "Price":
            f"{currency}{price:.2f}",

            "Qty":
            qty,

            "Total":
            f"{currency}{qty*price:.2f}"

        })



    return pd.DataFrame(quote)






# ==========================
# PDF CREATION
# ==========================

def create_pdf(
    quote,
    project_name,
    region
):


    output=BytesIO()


    doc=SimpleDocTemplate(
        output,
        pagesize=A4
    )


    styles=getSampleStyleSheet()

    elements=[]


    today=date.today().strftime(
        "%d/%m/%Y"
    )


    currency="£" if region=="Northern Ireland" else "€"



    elements.append(
        Paragraph(
            "<b>Quotation</b>",
            styles["Title"]
        )
    )


    elements.append(
        Spacer(1,15)
    )


    elements.append(
        Paragraph(
            f"<b>Project Name:</b> {project_name}",
            styles["Normal"]
        )
    )


    elements.append(
        Paragraph(
            "<b>Prepared By:</b> Quote Bot",
            styles["Normal"]
        )
    )


    elements.append(
        Paragraph(
            f"<b>Date:</b> {today}",
            styles["Normal"]
        )
    )


    elements.append(
        Spacer(1,20)
    )



    table_data=[

        [
            "Ref",
            "Product",
            "Code",
            "Type",
            "Warranty",
            "Price",
            "Qty",
            "Total"
        ]

    ]



    for _,x in quote.iterrows():


        table_data.append(

            [

                x["Ref"],
                x["Product"],
                x["Product Code"],
                x["Type"],
                x["Warranty"],
                x["Price"],
                x["Qty"],
                x["Total"]

            ]

        )



    table=Table(
        table_data,
        repeatRows=1
    )


    table.setStyle(

        TableStyle(

            [

                ("GRID",
                (0,0),
                (-1,-1),
                0.5,
                None)

            ]

        )

    )


    elements.append(table)


    elements.append(
        Spacer(1,20)
    )



    total=0


    for x in quote["Total"]:

        total += float(

            x.replace("€","")
             .replace("£","")

        )



    elements.append(

        Paragraph(

            f"<b>Total Project Price (Ex VAT): {currency}{total:.2f}</b>",

            styles["Normal"]

        )

    )



    doc.build(elements)


    output.seek(0)

    return output






# ==========================
# EXCEL
# ==========================

# ==========================
# EXCEL
# ==========================

def create_excel(quote):

    wb = load_workbook(
        BASE_DIR / "Quote Temp.xlsx"
    )

    ws = wb.active
    row = 2
    grand_total = 0

    for _, x in quote.iterrows():
        ws[f"A{row}"] = x["Ref"]
        ws[f"B{row}"] = x["Product"]
        ws[f"C{row}"] = x["Product Code"]
        ws[f"D{row}"] = x["Type"]
        ws[f"E{row}"] = x["Mode"]
        ws[f"F{row}"] = x["Mounting"]
        ws[f"G{row}"] = x["Colour"]
        ws[f"H{row}"] = x["IP"]
        ws[f"I{row}"] = x["Lumens"]
        ws[f"J{row}"] = x["Warranty"]

        price = float(
            x["Price"]
            .replace("€", "")
            .replace("£", "")
        )
        qty = int(x["Qty"])
        total_value = round(price * qty, 2)
        grand_total += total_value

        ws[f"K{row}"] = price
        ws[f"K{row}"].number_format = (
            '£#,##0.00' if "£" in x["Price"] else '€#,##0.00'
        )
        ws[f"L{row}"] = qty

        # Store the calculated value, not a formula.
        # openpyxl does not calculate Excel formulas itself.
        ws[f"M{row}"] = total_value
        ws[f"M{row}"].number_format = (
            '£#,##0.00' if "£" in x["Price"] else '€#,##0.00'
        )
        row += 1

    # Clear unused product rows.
    for r in range(row, 21):
        for c in range(1, 14):
            ws.cell(r, c).value = None

    ws["C21"] = "Commissioning"
    ws["K21"] = 0
    ws["M21"] = 0

    ws["C22"] = "Lux Levels"
    ws["K22"] = 0
    ws["M22"] = 0

    # Grand Total - actual value, not an unevaluated formula.
    ws["A23"] = "Total Project Price (Ex VAT)"
    ws["M23"] = grand_total

    last_currency = "€"
    if len(quote) and "£" in quote.iloc[0]["Price"]:
        last_currency = "£"

    ws["M23"].number_format = (
        '£#,##0.00' if last_currency == "£" else '€#,##0.00'
    )

    try:
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
        wb.calculation.calcMode = "auto"
    except Exception:
        pass

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output


def create_excel_file(quote):


    wb = load_workbook(
        BASE_DIR / "Quote Temp.xlsx"
    )

    ws = wb.active

    row = 2

    grand_total = 0


    for _, x in quote.iterrows():

        ws[f"A{row}"] = x["Ref"]
        ws[f"B{row}"] = x["Product"]
        ws[f"C{row}"] = x["Product Code"]
        ws[f"D{row}"] = x["Type"]
        ws[f"E{row}"] = x["Mode"]
        ws[f"F{row}"] = x["Mounting"]
        ws[f"G{row}"] = x["Colour"]
        ws[f"H{row}"] = x["IP"]
        ws[f"I{row}"] = x["Lumens"]
        ws[f"J{row}"] = x["Warranty"]


        price = float(
            x["Price"]
            .replace("€", "")
            .replace("£", "")
        )

        qty = int(x["Qty"])

        total_value = round(
            price * qty,
            2
        )

        grand_total += total_value


        ws[f"K{row}"] = price

        ws[f"K{row}"].number_format = (

            '£#,##0.00'

            if "£" in x["Price"]

            else '€#,##0.00'

        )


        ws[f"L{row}"] = qty


        # STORE ACTUAL VALUE (NO FORMULA)
        ws[f"M{row}"] = total_value

        ws[f"M{row}"].number_format = (

            '£#,##0.00'

            if "£" in x["Price"]

            else '€#,##0.00'

        )


        row += 1


    # Clear unused rows
    for r in range(row, 21):

        for c in range(1, 14):

            ws.cell(r, c).value = None


    # Commissioning row
    ws["C21"] = "Commissioning"
    ws["K21"] = 0
    ws["M21"] = 0

    # Lux row
    ws["C22"] = "Lux Levels"
    ws["K22"] = 0
    ws["M22"] = 0

    # Grand Total
    ws["A23"] = "Total Project Price (Ex VAT)"
    ws["M23"] = grand_total


    last_currency = "€"

    if len(quote):

        if "£" in quote.iloc[0]["Price"]:

            last_currency = "£"


    ws["M23"].number_format = (

        '£#,##0.00'

        if last_currency == "£"

        else '€#,##0.00'

    )


    excel_path = "/tmp/generated_quote.xlsx"

    wb.save(excel_path)

    return excel_path


def create_quote_image(excel_path, currency_symbol):

    from openpyxl import load_workbook
    from PIL import Image, ImageDraw, ImageFont

    wb = load_workbook(
        excel_path,
        data_only=True
    )

    ws = wb.active

    # ========================================================
    # HIGH RESOLUTION SCALE
    # ========================================================

    SCALE = 3

    cols = 13

    # Count actual product rows
    product_rows = 0

    for r in range(2, 21):

        if ws[f"B{r}"].value:
            product_rows += 1

    # Header + products + commissioning + lux + total
    rows = product_rows + 4

    cell_width = 110 * SCALE
    cell_height = 44 * SCALE

    img = Image.new(
        "RGB",
        (
            cols * cell_width,
            rows * cell_height
        ),
        "white"
    )

    draw = ImageDraw.Draw(img)

    # ========================================================
    # UNICODE FONT
    # ========================================================

    font_path = (
        "/usr/share/fonts/truetype/liberation/"
        "LiberationSans-Regular.ttf"
    )

    bold_font_path = (
        "/usr/share/fonts/truetype/liberation/"
        "LiberationSans-Bold.ttf"
    )

    try:

        font = ImageFont.truetype(
            font_path,
            22 * SCALE
        )

        total_font = ImageFont.truetype(
            bold_font_path,
            24 * SCALE
        )

    except Exception as e:

        print("Font loading error:", e)

        font = ImageFont.load_default()
        total_font = font

    # ========================================================
    # EXCEL ROWS
    # ========================================================

    excel_rows = [1]

    for r in range(2, 2 + product_rows):

        excel_rows.append(r)

    excel_rows.extend(
        [21, 22, 23]
    )

    image_row = 1

    # ========================================================
    # DRAW ROWS
    # ========================================================

    for excel_row in excel_rows:

        y1 = (
            image_row - 1
        ) * cell_height

        y2 = y1 + cell_height

        # ====================================================
        # TOTAL ROW
        # ====================================================

        if excel_row == 23:

            draw.rectangle(
                [
                    0,
                    y1,
                    12 * cell_width,
                    y2
                ],
                fill="#FFC000",
                outline="black",
                width=2 * SCALE
            )

            draw.text(
                (
                    10 * SCALE,
                    y1 + 12 * SCALE
                ),
                "Total Project Price (Ex VAT)",
                fill="black",
                font=total_font
            )

            draw.rectangle(
                [
                    12 * cell_width,
                    y1,
                    13 * cell_width,
                    y2
                ],
                fill="#00B0F0",
                outline="black",
                width=2 * SCALE
            )

            total_value = ws["M23"].value

            if total_value is None:
                total_value = 0

            total_text = (
                f"{currency_symbol}"
                f"{total_value:,.2f}"
            )

            draw.text(
                (
                    (12 * cell_width) + 5 * SCALE,
                    y1 + 12 * SCALE
                ),
                total_text,
                fill="white",
                font=total_font
            )

            image_row += 1

            continue

        # ====================================================
        # NORMAL ROWS
        # ====================================================

        for c in range(1, cols + 1):

            cell = ws.cell(
                excel_row,
                c
            )

            x1 = (
                c - 1
            ) * cell_width

            x2 = x1 + cell_width

            colour = "white"

            fill = str(
                cell.fill.fgColor.rgb
            )

            if "FFC000" in fill:

                colour = "#FFC000"

            elif "FFFF00" in fill:

                colour = "#FFFF00"

            elif "D9D9D9" in fill:

                colour = "#D9D9D9"

            elif "00B0F0" in fill:

                colour = "#00B0F0"

            draw.rectangle(
                [
                    x1,
                    y1,
                    x2,
                    y2
                ],
                fill=colour,
                outline="black",
                width=2 * SCALE
            )

            value = cell.value

            if value is None:

                value = ""

            # =================================================
            # CURRENCY COLUMNS
            # =================================================

            if (
                c in [11, 13]
                and isinstance(
                    value,
                    (int, float)
                )
            ):

                value = (
                    f"{currency_symbol}"
                    f"{value:,.2f}"
                )

            elif isinstance(
                value,
                float
            ):

                value = f"{value:,.2f}"

            elif isinstance(
                value,
                int
            ):

                value = str(value)

            draw.text(
                (
                    x1 + 4 * SCALE,
                    y1 + 12 * SCALE
                ),
                str(value),
                fill="black",
                font=font
            )

        image_row += 1

    # ========================================================
    # SAVE HIGH-RES IMAGE
    # ========================================================

    image_path = (
        "/tmp/quotation_page.png"
    )

    img.save(
        image_path,
        format="PNG"
    )

    return image_path

def create_pdf_quote(
    quote,
    project_name,
    region
):
    """
    Create the final quotation PDF.

    Page 1, 2 and 4 come from the existing Ventilux template.
    Page 3 is drawn as real PDF text/vector cells rather than a
    rasterized PNG. This keeps the text sharp and makes € / £
    render correctly using an embedded Unicode font.
    """

    from datetime import datetime
    from io import BytesIO

    # ==========================================================
    # CURRENCY
    # ==========================================================

    currency_symbol = (
        "£"
        if region == "Northern Ireland"
        else "€"
    )

    # ==========================================================
    # OPEN TEMPLATE
    # ==========================================================

    pdf = fitz.open(
        get_file("SAL-119-04.pdf")
    )

    # ==========================================================
    # PAGE 1 - COVER
    # ==========================================================

    page = pdf[0]

    today = datetime.today().strftime(
        "%d.%m.%Y"
    )

    # Remove existing values
    page.draw_rect(
        fitz.Rect(300, 655, 520, 680),
        fill=(0, 0, 0),
        color=(0, 0, 0)
    )

    page.draw_rect(
        fitz.Rect(300, 685, 560, 710),
        fill=(0, 0, 0),
        color=(0, 0, 0)
    )

    page.draw_rect(
        fitz.Rect(300, 715, 520, 740),
        fill=(0, 0, 0),
        color=(0, 0, 0)
    )

    page.insert_text(
        (305, 676),
        str(project_name),
        fontsize=10,
        color=(1, 1, 1)
    )

    page.insert_text(
        (305, 706),
        "Quote Bot",
        fontsize=10,
        color=(1, 1, 1)
    )

    page.insert_text(
        (305, 736),
        today,
        fontsize=10,
        color=(1, 1, 1)
    )

    # ==========================================================
    # PAGE 3 - REAL VECTOR/TEXT TABLE
    # ==========================================================

    page3 = pdf[2]

    # Use a Unicode font so both € and £ are embedded correctly.
    font_candidates = [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
    ]

    bold_font_candidates = [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
        "/usr/share/fonts/truetype/liberation2/LiberationSans-Bold.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
    ]

    font_path = next(
        (p for p in font_candidates if Path(p).exists()),
        None
    )

    bold_font_path = next(
        (p for p in bold_font_candidates if Path(p).exists()),
        None
    )

    if not font_path:
        raise FileNotFoundError(
            "A Unicode PDF font could not be found. "
            "DejaVu Sans or Liberation Sans is required "
            "to render € and £ correctly."
        )

    if not bold_font_path:
        bold_font_path = font_path

    # ==========================================================
    # TABLE DATA
    # ==========================================================

    headers = [
        "Ref",
        "Product",
        "Product Code",
        "Type",
        "Mode",
        "Mounting",
        "Colour",
        "IP",
        "Lumens",
        "Warranty",
        "Price",
        "Qty",
        "Total"
    ]

    rows = [headers]

    for _, x in quote.iterrows():

        rows.append([
            str(x.get("Ref", "")),
            str(x.get("Product", "")),
            str(x.get("Product Code", "")),
            str(x.get("Type", "")),
            str(x.get("Mode", "")),
            str(x.get("Mounting", "")),
            str(x.get("Colour", "")),
            str(x.get("IP", "")),
            str(x.get("Lumens", "")),
            str(x.get("Warranty", "")),
            str(x.get("Price", "")),
            str(x.get("Qty", "")),
            str(x.get("Total", ""))
        ])

    # ==========================================================
    # COMMISSIONING / LUX / TOTAL
    # ==========================================================

    commissioning_value = 0.0
    lux_value = 0.0

    rows.append([
        "",
        "",
        "Commissioning",
        "",
        "",
        "",
        "",
        "",
        "",
        "",
        f"{currency_symbol}{commissioning_value:,.2f}",
        "",
        f"{currency_symbol}{commissioning_value:,.2f}"
    ])

    rows.append([
        "",
        "",
        "Lux Levels",
        "",
        "",
        "",
        "",
        "",
        "",
        "",
        f"{currency_symbol}{lux_value:,.2f}",
        "",
        f"{currency_symbol}{lux_value:,.2f}"
    ])

    grand_total = 0.0

    for value in quote["Total"]:
        try:
            grand_total += float(
                str(value)
                .replace("€", "")
                .replace("£", "")
                .replace(",", "")
                .strip()
            )
        except (ValueError, TypeError):
            pass

    rows.append([
        "TOTAL",
        "",
        "Total Project Price (Ex VAT)",
        "",
        "",
        "",
        "",
        "",
        "",
        "",
        "",
        "",
        f"{currency_symbol}{grand_total:,.2f}"
    ])

    # ==========================================================
    # TABLE GEOMETRY
    # ==========================================================

    # A4 page width is ~595 points.
    # These widths total 575 points and leave a 10pt margin each side.
    col_widths = [
        22,   # Ref
        68,   # Product - wider so names stay on one line
        74,   # Product Code
        42,   # Type
        62,   # Mode - wider for "Non-Maintained"
        52,   # Mounting
        38,   # Colour
        20,   # IP
        44,   # Lumens
        35,   # Warranty
        42,   # Price
        28,   # Qty
        48    # Total
    ]

    table_x = 10
    table_y = 125

    header_height = 30
    normal_row_height = 28
    special_row_height = 28
    total_row_height = 32

    # ==========================================================
    # PDF FONT HELPERS
    # ==========================================================

    regular_font = fitz.Font(
        fontfile=font_path
    )

    bold_font = fitz.Font(
        fontfile=bold_font_path
    )

    def wrap_text(value, max_width, font_obj, fontsize):
        """
        Wrap text by measured PDF width, including long product codes
        that do not contain spaces.
        """

        value = str(value).strip()

        if not value:
            return [""]

        lines = []
        current = ""

        for char in value:

            candidate = current + char

            try:
                width = font_obj.text_length(
                    candidate,
                    fontsize=fontsize
                )
            except Exception:
                width = len(candidate) * fontsize * 0.55

            if current and width > max_width:
                lines.append(current)
                current = char
            else:
                current = candidate

        if current:
            lines.append(current)

        return lines or [""]

    def draw_cell(
        rect,
        value,
        font_obj,
        fontsize,
        fill,
        text_color=(0, 0, 0),
        align=0,
        bold=False
    ):
        page3.draw_rect(
            rect,
            fill=fill,
            color=(0, 0, 0),
            width=0.6
        )

        padding = 3

        text_rect = fitz.Rect(
            rect.x0 + padding,
            rect.y0 + padding,
            rect.x1 - padding,
            rect.y1 - padding
        )

        lines = wrap_text(
            value,
            max(5, text_rect.width),
            font_obj,
            fontsize
        )

        line_height = fontsize + 1.2
        max_lines = max(
            1,
            int(text_rect.height // line_height)
        )

        if len(lines) > max_lines:
            lines = lines[:max_lines]

            if lines:
                last = lines[-1]

                while last and font_obj.text_length(
                    last + "...",
                    fontsize=fontsize
                ) > text_rect.width:
                    last = last[:-1]

                lines[-1] = last + "..."

        total_height = len(lines) * line_height

        start_y = (
            text_rect.y0
            + max(
                0,
                (text_rect.height - total_height) / 2
            )
        )

        for line in lines:

            try:
                line_width = font_obj.text_length(
                    line,
                    fontsize=fontsize
                )
            except Exception:
                line_width = len(line) * fontsize * 0.55

            if align == 1:
                x = (
                    text_rect.x0
                    + (text_rect.width - line_width) / 2
                )
            elif align == 2:
                x = (
                    text_rect.x1
                    - line_width
                )
            else:
                x = text_rect.x0

            page3.insert_text(
                (x, start_y + fontsize),
                line,
                fontsize=fontsize,
                fontname="DejaVuSans",
                fontfile=(
                    bold_font_path
                    if bold
                    else font_path
                ),
                color=text_color
            )

            start_y += line_height

    # ==========================================================
    # DRAW TABLE
    # ==========================================================

    y = table_y

    for row_index, row_data in enumerate(rows):

        if row_index == 0:

            row_height = header_height
            fontsize = 6.8
            row_fill = (1.0, 0.67, 0.0)
            font_obj = bold_font

        elif row_index == len(rows) - 1:

            row_height = total_row_height
            fontsize = 7.4
            row_fill = (1.0, 0.67, 0.0)
            font_obj = bold_font

        elif row_index >= len(rows) - 3:

            row_height = special_row_height
            fontsize = 6.8
            row_fill = (1.0, 1.0, 0.0)
            font_obj = regular_font

        else:

            row_height = normal_row_height
            fontsize = 6.8
            row_fill = (1.0, 1.0, 1.0)
            font_obj = regular_font

        if row_index == len(rows) - 1:

            merged_width = sum(col_widths[:12])

            label_rect = fitz.Rect(
                table_x,
                y,
                table_x + merged_width,
                y + row_height
            )

            draw_cell(
                label_rect,
                "Total Project Price (Ex VAT)",
                bold_font,
                7.2,
                (1.0, 0.67, 0.0),
                align=1,
                bold=True
            )

            total_rect = fitz.Rect(
                table_x + merged_width,
                y,
                table_x + merged_width + col_widths[12],
                y + row_height
            )

            draw_cell(
                total_rect,
                row_data[12],
                bold_font,
                7.4,
                (0.0, 0.69, 0.94),
                text_color=(1, 1, 1),
                align=2,
                bold=True
            )

        else:

            x = table_x

            for col_index, value in enumerate(row_data):

                width = col_widths[col_index]

                rect = fitz.Rect(
                    x,
                    y,
                    x + width,
                    y + row_height
                )

                # Currency / quantity / short numeric fields centered.
                centered = col_index in [
                    0, 7, 8, 9, 10, 11, 12
                ]

                draw_cell(
                    rect,
                    value,
                    font_obj,
                    fontsize,
                    row_fill,
                    align=1 if centered else 0,
                    bold=(row_index == 0)
                )

                x += width

        y += row_height

    # ==========================================================
    # SAVE
    # ==========================================================

    temp_pdf = "/tmp/final_quote.pdf"

    pdf.save(
        temp_pdf,
        garbage=4,
        deflate=True
    )

    pdf.close()

    with open(
        temp_pdf,
        "rb"
    ) as f:

        pdf_bytes = BytesIO(
            f.read()
        )

    pdf_bytes.seek(0)

    return pdf_bytes

# ==========================
# UI
# ==========================


st.title(
    "💡 Product Quotation Bot"
)



project_name=st.text_input(
    "Project Name"
)



region=st.radio(
    "Region",
    [
        "Ireland",
        "Northern Ireland"
    ]
)



customer_type=st.radio(
    "Customer Type",
    [
        "Contractor",
        "Wholesaler",
        "OPW"
    ]
)



custom=st.radio(
    "Custom Reference?",
    [
        "No",
        "Yes"
    ]
)



output_type=st.radio(
    "Quotation Output",
    [
        "Excel",
        "PDF",
        "Both"
    ]
)



products=st.text_area(
    "Products",
    height=250
)



warranty=st.radio(
    "Warranty",
    [
        "4 Years",
        "8 Years"
    ]
)





if st.button(
    "Generate Quotation"
):


    lines=[

        x.strip()
        for x in products.split("\n")
        if x.strip()

    ]



    items=[]

    refs={}



    for line in lines:


        if custom=="Yes" and "|" in line:


            r,p=line.split(
                "|",
                1
            )


            p=p.strip()


            refs[p]=r.strip()


            items.append(p)


        else:

            items.append(line)





    quote=create_quote(

        items,

        customer_type,

        warranty,

        refs,

        region

    )



    st.subheader(
        "Quotation"
    )


    st.dataframe(

        quote,

        hide_index=True,

        use_container_width=True

    )



    if len(quote):


        total=0


        for x in quote["Total"]:

            total += float(

                x.replace("€","")
                 .replace("£","")

            )



        symbol="£" if region=="Northern Ireland" else "€"



        st.success(

            f"Grand Total: {symbol}{total:.2f}"

        )



        # =========================================================================================================================
        # EXCEL DOWNLOAD
        # =========================================================================================================================

        if output_type in ["Excel","Both"]:


            excel=create_excel(
                quote
            )


            st.download_button(

                "⬇️ Download Excel Quotation",

                excel,

                file_name="Quotation.xlsx",

                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

            )



        # ==========================
        # PDF DOWNLOAD
        # ==========================

        if output_type in ["PDF","Both"]:


            pdf=create_pdf_quote(

                quote,

                project_name,

                region

            )


            st.download_button(

                "📄 Download PDF Quotation",

                pdf,

                file_name=f"{project_name}.pdf",

                mime="application/pdf"

            )


    else:


        st.error(
            "No matching products found. Check product name/code."
        )
