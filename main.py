import dash
from dash import dcc, html, Input, Output, State
import dash_bootstrap_components as dbc
from flask import send_file
import os
from mom import replace_placeholders  # Import the function from your other file

# Initialize the Dash app
app = dash.Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP])

# Layout for the form
app.layout = dbc.Container(
    [
        html.H1("MOM Generator", className="text-center mt-4"),
        dbc.Form(
            [
                dbc.Label("Department", html_for="dept"),
                dbc.Input(id="dept", placeholder="Enter Department", type="text", required=True, className="mb-3"),
                
                dbc.Label("Date", html_for="date"),
                dbc.Input(id="date", type="date", required=True, className="mb-3"),
                
                dbc.Label("Start Time", html_for="start_time"),
                dbc.Input(id="start_time", type="time", required=True, className="mb-3"),
                
                dbc.Label("End Time", html_for="end_time"),
                dbc.Input(id="end_time", type="time", required=True, className="mb-3"),
                
                dbc.Label("Agenda", html_for="agenda"),
                dbc.Input(id="agenda", placeholder="Enter Agenda", required=True, className="mb-3"),
                
                dbc.Label("Absent Members (comma-separated)", html_for="abs"),
                dbc.Input(id="abs", placeholder="Enter names, e.g., Anurag, Tarun, Akshat", type="text", required=True, className="mb-3"),
                
                dbc.Label("Additional Notes (Bullet Points)", html_for="notes"),
                dbc.Textarea(id="notes", placeholder="Enter bullet points, each on a new line", rows=4, required=True, className="mb-3"),
                
                dbc.Row(
                    [
                        dbc.Col(
                            dbc.Button("Generate PDF", id="generate-btn", color="primary", className="mt-3"),
                            width="auto",
                        ),
                        dbc.Col(
                            html.Div(id="download-link", className="mt-3"),
                            width="auto",
                        ),
                    ],
                    justify="start",
                    align="center",
                ),
            ],
            className="p-4 border rounded",
        ),
    ],
    className="mt-5",
)

# Callback to generate and download PDF
@app.callback(
    Output("download-link", "children"),
    Input("generate-btn", "n_clicks"),
    State("dept", "value"),
    State("date", "value"),
    State("start_time", "value"),
    State("end_time", "value"),
    State("agenda", "value"),
    State("abs", "value"),
    State("notes", "value"),
    prevent_initial_call=True,
)
def generate_pdf(n_clicks, dept, date, start_time, end_time, agenda, abs_list, notes):
    if not (dept and date and start_time and end_time and agenda and abs_list and notes):
        return dbc.Alert("Please fill in all fields.", color="danger")

    # Placeholder replacements
    replacements = {
        "{dept}": dept,
        "{date}": date,
        "{start_time}": start_time,
        "{end_time}": end_time,
        "{agenda}": agenda,
        "{abs}": [name.strip() for name in abs_list.split(",")],
        "{_}": notes.split("\n"),
    }

    # Paths for input and output files
    template_path = "template.docx"
    pdf_output_path = "output.pdf"

    # Generate the PDF using the external function
    replace_placeholders(template_path, pdf_output_path, replacements)

    # Create a download link
    return dcc.Link(
        "Download PDF",
        href="/download-pdf",
        target="_blank",
        className="btn btn-success mt-3",
    )

# Route for downloading the PDF
@app.server.route("/download-pdf")
def download_pdf():
    pdf_path = "output.pdf"
    if os.path.exists(pdf_path):
        return send_file(pdf_path, as_attachment=True)
    else:
        return "PDF not found", 404

if __name__ == "__main__":
    app.run_server(debug=True)
