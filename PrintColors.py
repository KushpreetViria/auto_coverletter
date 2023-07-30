class ColorPrint:
    COLORS = {
        "black": "30",
        "red": "31",
        "green": "32",
        "yellow": "33",
        "blue": "34",
        "magenta": "35",
        "cyan": "36",
        "white": "37",
    }

    STYLES = {"normal": "0", "bold": "1", "italic": "3", "underline": "4", "negative": "7"}

    @classmethod
    def print_color(cls, text:str, color:str="normal", styles:list=None):
        if styles == None:
            styles = ["normal"]

        color_code = cls.COLORS.get(color.lower(), "37")
        style_code = ";".join([cls.STYLES.get(style.lower(), "0") for style in styles])
        print(f"\033[{style_code};{color_code}m{text}\033[0m")

    @classmethod
    def print_error(cls, text:str):
        cls.print_color(text, color="red", styles=["bold"])
