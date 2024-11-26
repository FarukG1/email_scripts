"""This module reads html files and replaces variables with values"""


class HTML:
    """Reads html files and replaces variables with values"""

    @staticmethod
    def read_html_file(file_path, variables):
        """Reads html file and replaces variables with values"""
        with open(file_path, "r", encoding="utf-8") as file:
            html = file.read()
            for key, value in variables.items():
                html = html.replace("{{" + key + "}}", value)
            title = html.split("<title>")[1].split("</title>")[0]
        return title, html
