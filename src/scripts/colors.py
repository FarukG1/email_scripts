"""Colors.py - Colors class for terminal output."""

import os


class Colors:
    """Colors class for terminal output."""

    purple = "\033[95m"
    blue = "\033[94m"
    cyan = "\033[96m"
    green = "\033[92m"
    warning = "\033[93m"
    error = "\033[91m"
    reset = "\033[0m"
    bold = "\033[1m"
    underline = "\033[4m"

    __dir__ = {
        "purple": "\033[95m",
        "blue": "\033[94m",
        "cyan": "\033[96m",
        "green": "\033[92m",
        "warning": "\033[93m",
        "error": "\033[91m",
        "reset": "\033[0m",
        "bold": "\033[1m",
        "underline": "\033[4m",
    }

    @classmethod
    def print(cls, color, text):
        """Print with color."""
        os.system("color")
        return print(f"{cls.__dir__[color]}{text}{cls.reset}")

    @classmethod
    def colored(cls, color, text):
        """String with color."""
        os.system("color")
        return f"{cls.__dir__[color]}{text}{cls.reset}"
