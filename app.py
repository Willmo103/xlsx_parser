import io
from tempfile import NamedTemporaryFile
from flask import Flask, render_template, request, send_file
import os
from amazon_report_parse import parse_amazon_report

# Initialize the Flask application
app = Flask(__name__)


# Initialize the templates directory
app.template_folder = "templates"


def save_uploaded_file_and_return_temp_path(uploaded_file):
    """
    Save the uploaded file to a temporary location and return its path.
    """
    temp_file = NamedTemporaryFile(delete=False, suffix=".xlsx")
    uploaded_file.save(temp_file.name)
    return temp_file.name


@app.route("/", methods=["GET", "POST"])
def index():
    """
    Flask endpoint for the main page.
    Renders the upload page on GET request.
    Processes the uploaded file and returns the processed file as a download on POST request.
    """
    if request.method == "POST":
        file = request.files["file"]
        if file:
            input_temp_path = save_uploaded_file_and_return_temp_path(file)
            output_temp_file = NamedTemporaryFile(delete=False, suffix=".xlsx")
            output_temp_path = output_temp_file.name
            output_temp_file.close()

            parse_amazon_report(input_temp_path, output_temp_path)

            # Serve the file and then delete temporary files
            with open(output_temp_path, "rb") as f:
                content = f.read()
            os.remove(input_temp_path)
            os.remove(output_temp_path)
            return send_file(
                io.BytesIO(content),
                as_attachment=True,
                download_name="processed_output.xlsx",
                mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    # If GET request, render the upload page
    return render_template("uploads.html")


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=9876)
