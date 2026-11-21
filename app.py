from flask import Flask, request, redirect, render_template, send_file
import sqlite3
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from io import BytesIO

app = Flask(__name__)


def get_database():
    connection = sqlite3.connect("yeastbank.db")
    connection.row_factory = sqlite3.Row
    return connection

@app.route("/")
def home():
    return render_template("home.html")

@app.route("/strains")
def strains():

    connection = get_database()

    search = request.args.get("search", "").strip()
    selected_species = request.args.get("species", "").strip()
    selected_application = request.args.get("application", "").strip()

    query = "SELECT * FROM strains WHERE 1=1"
    parameters = []

    if search:
        query += " AND (strain_code LIKE ? OR strain_name LIKE ?)"
        parameters.extend([f"%{search}%", f"%{search}%"])

    if selected_species:
        query += " AND species = ?"
        parameters.append(selected_species)

    if selected_application:
        query += " AND application = ?"
        parameters.append(selected_application)

    query += " ORDER BY id"

    records = connection.execute(
        query, parameters
    ).fetchall()

    species_list = connection.execute(
        "SELECT DISTINCT species FROM strains WHERE species IS NOT NULL AND species != '' ORDER BY species"
    ).fetchall()

    application_list = connection.execute(
        "SELECT DISTINCT application FROM strains WHERE application IS NOT NULL AND application != '' ORDER BY application"
    ).fetchall()

    connection.close()

    return render_template(
        "strains.html",
        strains=records,
        search=search,
        species_list=[row["species"] for row in species_list],
        application_list=[row["application"] for row in application_list],
        selected_species=selected_species,
        selected_application=selected_application
    )

    html = """
    <h1>YeastBank Strain Database</h1>
    <a href="/">Home</a> |
    <a href="/add-strain">Add New Strain</a>
    <table border="1" cellpadding="10">
        <tr>
            <th>ID</th>
            <th>Strain Code</th>
            <th>Strain Name</th>
            <th>Species</th>
            <th>Application</th>
        </tr>
    """

    for record in records:
        html += f"""
        <tr>
            <td>{record['id']}</td>
            <td>{record['strain_code']}</td>
            <td>{record['strain_name']}</td>
            <td>{record['species']}</td>
            <td>{record['application']}</td>
        </tr>
        """

    html += "</table>"
    html += f"<p>Total strains: {len(records)}</p>"

    return html
@app.route("/add-strain", methods=["GET", "POST"])
def add_strain():
    if request.method == "POST":

        strain_code = request.form["strain_code"].strip()
        strain_name = request.form["strain_name"].strip()
        species = request.form["species"].strip()
        application = request.form["application"].strip()

        if not all([strain_code, strain_name, species, application]):
            return "Please fill in all fields."

        connection = get_database()

        try:
            connection.execute(
                """
                INSERT INTO strains
                (strain_code, strain_name, species, application)
                VALUES (?, ?, ?, ?)
                """,
                (strain_code, strain_name, species, application)
            )

            connection.commit()

        except sqlite3.IntegrityError:
            connection.close()
            return "That strain code already exists. Please use a different code."

        connection.close()

        return redirect("/strains")

    return render_template("add_strain.html")


    return """
    <h1>Add a New Yeast Strain</h1>

    <form method="POST">
        <label>Strain Code:</label><br>
        <input type="text" name="strain_code" required><br><br>

        <label>Strain Name:</label><br>
        <input type="text" name="strain_name" required><br><br>

        <label>Species:</label><br>
        <input type="text" name="species" required><br><br>

        <label>Application:</label><br>
        <input type="text" name="application" required><br><br>

        <button type="submit">Save Strain</button>
    </form>

    <br>
    <a href="/strains">Back to Strain Database</a>
    """
@app.route("/performance", methods=["GET", "POST"])
def performance():

    connection = get_database()

    if request.method == "POST":

        strain_id = request.form["strain_id"]

        ethanol = request.form.get("ethanol_tolerance", "").strip()
        temperature = request.form.get("temperature_range", "").strip()
        flocculation = request.form.get("flocculation", "").strip()
        attenuation = request.form.get("attenuation", "").strip()
        aroma = request.form.get("aroma_profile", "").strip()

        connection.execute("""
            UPDATE strains
            SET ethanol_tolerance = ?,
                temperature_range = ?,
                flocculation = ?,
                attenuation = ?,
                aroma_profile = ?
            WHERE id = ?
        """, (
            ethanol,
            temperature,
            flocculation,
            attenuation,
            aroma,
            strain_id
        ))

        connection.commit()
        connection.close()

        return redirect("/strains")

    records = connection.execute(
        "SELECT id, strain_code FROM strains ORDER BY id"
    ).fetchall()

    connection.close()

    return render_template(
        "performance.html",
        strains=records
    )
@app.route("/strain/<int:strain_id>")
def strain_detail(strain_id):

    connection = get_database()

    strain = connection.execute(
        "SELECT * FROM strains WHERE id = ?",
        (strain_id,)
    ).fetchone()

    connection.close()

    if strain is None:
        return "Yeast strain not found", 404

    return render_template(
        "strain_detail.html",
        strain=strain
    )
@app.route("/edit-strain/<int:strain_id>", methods=["GET", "POST"])
def edit_strain(strain_id):

    connection = get_database()

    strain = connection.execute(
        "SELECT * FROM strains WHERE id = ?",
        (strain_id,)
    ).fetchone()

    if strain is None:
        connection.close()
        return "Yeast strain not found", 404

    if request.method == "POST":

        strain_code = request.form["strain_code"].strip()
        strain_name = request.form["strain_name"].strip()
        species = request.form["species"].strip()
        application = request.form["application"].strip()

        if not all([strain_code, strain_name, species, application]):
            connection.close()
            return "Please fill in all fields."

        try:
            connection.execute("""
                UPDATE strains
                SET strain_code = ?,
                    strain_name = ?,
                    species = ?,
                    application = ?
                WHERE id = ?
            """, (
                strain_code,
                strain_name,
                species,
                application,
                strain_id
            ))

            connection.commit()

        except sqlite3.IntegrityError:
            connection.close()
            return "This strain code already exists."

        connection.close()

        return redirect("/strains")

    connection.close()

    return render_template(
        "edit_strain.html",
        strain=strain
    )
@app.route("/delete-strain/<int:strain_id>", methods=["POST"])
def delete_strain(strain_id):

    connection = get_database()

    strain = connection.execute(
        "SELECT id FROM strains WHERE id = ?",
        (strain_id,)
    ).fetchone()

    if strain is None:
        connection.close()
        return "Yeast strain not found", 404

    connection.execute(
        "UPDATE aroma_compounds SET strain_id = NULL WHERE strain_id = ?",
        (strain_id,)
    )

    connection.execute(
        "DELETE FROM strains WHERE id = ?",
        (strain_id,)
    )

    connection.commit()
    connection.close()

    return redirect("/strains")

@app.route("/aroma")
def aroma():

    connection = get_database()

    search = request.args.get("search", "").strip()

    compounds = connection.execute("""
        SELECT
            aroma_compounds.*,
            strains.strain_code
        FROM aroma_compounds
        LEFT JOIN strains
        ON aroma_compounds.strain_id = strains.id
        WHERE aroma_compounds.compound_name LIKE ?
        ORDER BY aroma_compounds.compound_name
    """, (f"%{search}%",)).fetchall()

    connection.close()

    return render_template(
        "aroma.html",
        compounds=compounds,
        search=search
    )
@app.route("/add-aroma", methods=["GET", "POST"])
def add_aroma():

    connection = get_database()

    if request.method == "POST":

        compound_name = request.form.get(
            "compound_name", ""
        ).strip()

        chemical_category = request.form.get(
            "chemical_category", ""
        ).strip()

        chemical_formula = request.form.get(
            "chemical_formula", ""
        ).strip()

        aroma_description = request.form.get(
            "aroma_description", ""
        ).strip()

        strain_id = request.form.get("strain_id") or None

        if not compound_name:
            connection.close()
            return "Please enter a compound name.", 400

        connection.execute("""
            INSERT INTO aroma_compounds
            (compound_name, chemical_category,
             chemical_formula, aroma_description, strain_id)
            VALUES (?, ?, ?, ?, ?)
        """, (
            compound_name,
            chemical_category,
            chemical_formula,
            aroma_description,
            strain_id
        ))

        connection.commit()
        connection.close()

        return redirect("/aroma")

    strains = connection.execute(
        "SELECT id, strain_code FROM strains ORDER BY strain_code"
    ).fetchall()

    connection.close()

    return render_template(
        "add_aroma.html",
        strains=strains
    )
@app.route("/edit-aroma/<int:compound_id>", methods=["GET", "POST"])
def edit_aroma(compound_id):

    connection = get_database()

    compound = connection.execute(
        "SELECT * FROM aroma_compounds WHERE id = ?",
        (compound_id,)
    ).fetchone()

    if compound is None:
        connection.close()
        return "Aroma compound not found", 404

    if request.method == "POST":

        compound_name = request.form["compound_name"].strip()
        chemical_category = request.form["chemical_category"].strip()
        chemical_formula = request.form["chemical_formula"].strip()
        aroma_description = request.form["aroma_description"].strip()

        connection.execute("""
            UPDATE aroma_compounds
            SET compound_name = ?,
                chemical_category = ?,
                chemical_formula = ?,
                aroma_description = ?
            WHERE id = ?
        """, (
            compound_name,
            chemical_category,
            chemical_formula,
            aroma_description,
            compound_id
        ))

        connection.commit()
        connection.close()

        return redirect("/aroma")

    connection.close()

    return render_template(
        "edit_aroma.html",
        compound=compound
    )
@app.route("/delete-aroma/<int:compound_id>", methods=["POST"])
def delete_aroma(compound_id):

    connection = get_database()

    compound = connection.execute(
        "SELECT id FROM aroma_compounds WHERE id = ?",
        (compound_id,)
    ).fetchone()

    if compound is None:
        connection.close()
        return "Aroma compound not found", 404

    connection.execute(
        "DELETE FROM aroma_compounds WHERE id = ?",
        (compound_id,)
    )

    connection.commit()
    connection.close()

    return redirect("/aroma")
@app.route("/link-aroma", methods=["GET", "POST"])
def link_aroma():

    connection = get_database()

    if request.method == "POST":

        compound_id = request.form["compound_id"]
        strain_id = request.form["strain_id"]

        compound = connection.execute(
            "SELECT id FROM aroma_compounds WHERE id = ?",
            (compound_id,)
        ).fetchone()

        strain = connection.execute(
            "SELECT id FROM strains WHERE id = ?",
            (strain_id,)
        ).fetchone()

        if compound is None or strain is None:
            connection.close()
            return "Invalid compound or strain selection.", 400

        connection.execute("""
            UPDATE aroma_compounds
            SET strain_id = ?
            WHERE id = ?
        """, (strain_id, compound_id))

        connection.commit()
        connection.close()

        return redirect("/aroma")

    compounds = connection.execute(
        "SELECT id, compound_name FROM aroma_compounds ORDER BY compound_name"
    ).fetchall()

    strains = connection.execute(
        "SELECT id, strain_code FROM strains ORDER BY strain_code"
    ).fetchall()

    connection.close()

    return render_template(
        "link_aroma.html",
        compounds=compounds,
        strains=strains
    )
@app.route("/compare", methods=["GET", "POST"])
def compare():

    connection = get_database()

    strains = connection.execute(
        "SELECT * FROM strains ORDER BY strain_code"
    ).fetchall()

    comparison = None
    aroma_comparison = []
    message = ""

    if request.method == "POST":

        first_id = request.form.get("strain1")
        second_id = request.form.get("strain2")

        if not first_id or not second_id:
            message = "Please select two strains."

        elif first_id == second_id:
            message = "Please select two different strains."

        else:

            first = connection.execute(
                "SELECT * FROM strains WHERE id = ?",
                (first_id,)
            ).fetchone()

            second = connection.execute(
                "SELECT * FROM strains WHERE id = ?",
                (second_id,)
            ).fetchone()

            if first is None or second is None:

                message = "Invalid strain selection."

            else:

                comparison = [first, second]

                # Retrieve aroma compounds for both strains

                for strain in comparison:

                    aromas = connection.execute("""
                        SELECT compound_name, chemical_category, chemical_formula,
       aroma_description
FROM aroma_compounds
WHERE strain_id = ?
ORDER BY compound_name
                    """, (strain["id"],)).fetchall()

                    aroma_comparison.append(aromas)

    connection.close()

    return render_template(
        "compare.html",
        strains=strains,
        comparison=comparison,
        message=message,
        aroma_comparison=aroma_comparison
    )
# ==============================
# EXCEL EXPORT
# ==============================

@app.route("/export-excel")
def export_excel():

    connection = get_database()

    workbook = Workbook()

    # Summary Sheet
    summary = workbook.active
    summary.title = "Summary"

    summary.append(["YEASTBANK DATABASE SUMMARY"])
    summary.append(["Metric", "Total"])

    total_strains = connection.execute(
        "SELECT COUNT(*) FROM strains"
    ).fetchone()[0]

    total_compounds = connection.execute(
        "SELECT COUNT(*) FROM aroma_compounds"
    ).fetchone()[0]

    total_species = connection.execute(
        "SELECT COUNT(DISTINCT species) FROM strains WHERE species IS NOT NULL AND species != ''"
    ).fetchone()[0]

    total_applications = connection.execute(
        "SELECT COUNT(DISTINCT application) FROM strains WHERE application IS NOT NULL AND application != ''"
    ).fetchone()[0]

    summary.append(["Total Yeast Strains", total_strains])
    summary.append(["Total Aroma Compounds", total_compounds])
    summary.append(["Total Species", total_species])
    summary.append(["Total Applications", total_applications])

    summary.column_dimensions["A"].width = 30
    summary.column_dimensions["B"].width = 18

    # Summary Formatting
    for cell in summary[1]:
        cell.font = Font(bold=True, color="FFFFFF", size=14)
        cell.fill = PatternFill(
            fill_type="solid",
            fgColor="145A32"
        )

    for cell in summary[2]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(
            fill_type="solid",
            fgColor="228B22"
        )

    # Yeast Strains Sheet
    sheet1 = workbook.create_sheet("Yeast Strains")

    strains = connection.execute(
        "SELECT * FROM strains"
    ).fetchall()

    if strains:
        sheet1.append(list(strains[0].keys()))

        for row in strains:
            sheet1.append(list(row))

    # Aroma Compounds Sheet
    sheet2 = workbook.create_sheet("Aroma Compounds")

    aromas = connection.execute(
        "SELECT * FROM aroma_compounds"
    ).fetchall()

    if aromas:
        sheet2.append(list(aromas[0].keys()))

        for row in aromas:
            sheet2.append(list(row))

    # Professional Formatting
    for sheet in workbook.worksheets:

        for cell in sheet[1]:
            cell.font = Font(
                bold=True,
                color="FFFFFF",
                size=11
            )

            cell.fill = PatternFill(
                fill_type="solid",
                fgColor="145A32"
            )

            cell.alignment = Alignment(
                horizontal="center"
            )

        for row in sheet.iter_rows():
            for cell in row:
                cell.border = Border(
                    bottom=Side(
                        style="thin",
                        color="D5D8DC"
                    )
                )

        for column in sheet.columns:

            max_length = 0
            column_letter = column[0].column_letter

            for cell in column:
                if cell.value is not None:
                    max_length = max(
                        max_length,
                        len(str(cell.value))
                    )

            sheet.column_dimensions[column_letter].width = min(
                max_length + 4, 45
            )

        sheet.freeze_panes = "A2"

        if sheet.title != "Summary":
            sheet.auto_filter.ref = sheet.dimensions

        sheet.row_dimensions[1].height = 25

    connection.close()

    output = BytesIO()

    workbook.save(output)

    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name="YeastBank_Data.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


# ==============================
# DASHBOARD
# ==============================

@app.route("/dashboard")
def dashboard():

    connection = get_database()

    total_strains = connection.execute(
        "SELECT COUNT(*) FROM strains"
    ).fetchone()[0]

    total_compounds = connection.execute(
        "SELECT COUNT(*) FROM aroma_compounds"
    ).fetchone()[0]

    total_species = connection.execute(
        """SELECT COUNT(DISTINCT species)
           FROM strains
           WHERE species IS NOT NULL AND species != ''"""
    ).fetchone()[0]

    total_applications = connection.execute(
        """SELECT COUNT(DISTINCT application)
           FROM strains
           WHERE application IS NOT NULL AND application != ''"""
    ).fetchone()[0]

    species_data = connection.execute("""
        SELECT species, COUNT(*) AS total
        FROM strains
        WHERE species IS NOT NULL AND species != ''
        GROUP BY species
        ORDER BY species
    """).fetchall()

    species_labels = [
        row["species"] for row in species_data
    ]

    species_values = [
        row["total"] for row in species_data
    ]

    aroma_data = connection.execute("""
        SELECT chemical_category, COUNT(*) AS total
        FROM aroma_compounds
        WHERE chemical_category IS NOT NULL
        AND chemical_category != ''
        GROUP BY chemical_category
        ORDER BY chemical_category
    """).fetchall()

    aroma_labels = [
        row["chemical_category"] for row in aroma_data
    ]

    aroma_values = [
        row["total"] for row in aroma_data
    ]

    connection.close()

    return render_template(
        "dashboard.html",
        total_strains=total_strains,
        total_compounds=total_compounds,
        total_species=total_species,
        total_applications=total_applications,
        species_labels=species_labels,
        species_values=species_values,
        aroma_labels=aroma_labels,
        aroma_values=aroma_values
    )
if __name__ == "__main__":
    connection = get_database()

    print("EXACT AROMA TABLE COLUMN NAMES:")

    columns = connection.execute(
        "PRAGMA table_info(aroma_compounds)"
    ).fetchall()

    for column in columns:
        print(repr(column["name"]))

    connection.close()

    app.run(debug=True)