from datetime import datetime

import requests
import xlsxwriter
from bs4 import BeautifulSoup #scrape chart data from the wiki

class ArcaeaSong:
    def __init__(self, title: str, difficulty: str, constant: float, **kwargs):
        self.title = title
        self.difficulty = difficulty
        self.constant = constant
        self.artist = kwargs.get("artist")
        self.pack = kwargs.get("pack")

    def __str__(self):
        return f"{self.__class__}(title={self.title}, difficulty={self.difficulty}, constant={self.constant})"

    def __repr__(self):
        return f"{self.__class__}(title={self.title}, difficulty={self.difficulty}, constant={self.constant})"

    @property
    def difficulty_color(self):
        if self.difficulty == "BYD":
            return "#b22222"
        elif self.difficulty == "ETR":
            return "#6a5acd"
        elif self.difficulty == "FTR":
            return "#c71585"
        elif self.difficulty == "PRS":
            return "#32cd32"
        elif self.difficulty == "PST":
            return "#00bfff"

def scrape_song_list(variant: str = "8+") -> "list[ArcaeaSong]":
    if variant == "8+":
        url = "https://wikiwiki.jp/arcaea/%E8%AD%9C%E9%9D%A2%E5%AE%9A%E6%95%B0%E8%A1%A8"
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36"
        }
        with requests.get(url, headers=headers) as resp:
            if not resp.ok:
                raise Exception("Could not download webpage")
            soup = BeautifulSoup(resp.text, "html.parser")

        songs = []
        for row in soup.select("table > tbody > tr"):
            cells = row.find_all("td")
            if len(cells) < 7:
                continue

            difficulty = "PST"
            diff_color = cells[5]["style"].split("background-color:")[1].split(";")[0]
            if diff_color == "Firebrick":
                difficulty = "BYD"
            if diff_color == "Slateblue":
                difficulty = "ETR"
            if diff_color == "Mediumvioletred":
                difficulty = "FTR"
            if diff_color == "Mediumseagreen":
                difficulty = "PRS"

            songs.append(
                ArcaeaSong(
                    cells[2].get_text(),
                    difficulty,
                    float(cells[6].get_text()),
                    artist=cells[3].get_text(),
                    pack=cells[4].get_text(),
                )
            )

        return songs
    else:
        url = "https://wikiwiki.jp/arcaea/%E8%AD%9C%E9%9D%A2%E5%AE%9A%E6%95%B0%E8%A1%A8/%E8%AD%9C%E9%9D%A2%E5%AE%9A%E6%95%B0%E8%A1%A8%20%28Level%207%E4%BB%A5%E4%B8%8B%29"
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36"
        }
        with requests.get(url, headers=headers) as resp:
            if not resp.ok:
                raise Exception("Could not download webpage")
            soup = BeautifulSoup(resp.text, "html.parser")

        songs = []
        for row in soup.select("table > tbody > tr"):
            cells = row.find_all("td")
            if len(cells) < 5:
                continue

            difficulty = "PST"
            diff_color = cells[3]["style"].split("background-color:")[1].split(";")[0]
            if diff_color == "Firebrick":
                difficulty = "BYD"
            if diff_color == "Slateblue":
                difficulty = "ETR"
            if diff_color == "Mediumvioletred":
                difficulty = "FTR"
            if diff_color == "Mediumseagreen":
                difficulty = "PRS"

            songs.append(
                ArcaeaSong(
                    cells[0].get_text(),
                    difficulty,
                    float(cells[4].get_text()),
                    artist=cells[2].get_text(),
                )
            )
        return songs


def construct_workbook(songs: "list[ArcaeaSong]", output: str = "result.xlsx"):
    workbook = xlsxwriter.Workbook(output)
    worksheet = workbook.add_worksheet("SCORES")

    bold = workbook.add_format({"bold": True})
    worksheet.write(
        "A1",
        "Title",
        workbook.add_format({"bold": True, "bottom": 5}),
    )
    worksheet.set_column("A:A", 35)
    worksheet.write(
        "B1",
        "DIFF",
        workbook.add_format({"bold": True, "bottom": 5}),
    )
    worksheet.set_column("B:B", 4.35)
    worksheet.write(
        "C1",
        "CC",
        workbook.add_format({"bold": True, "bottom": 5}),
    )
    worksheet.set_column("C:C", 4.35)
    worksheet.write("D1", "SCORE", bold)
    worksheet.set_column("D:D", 14)
    worksheet.write("E1", "PLAY RATING", bold)
    worksheet.set_column("E:E", 14)

    cell_colors = ["#d9d2e9", "#cfe2f3"]
    cell_style = 0
    prev_constant = 0.0
    trailing_fragment = 0
    for (idx, song) in enumerate(songs):
        index = idx + 2
        if not song.constant == prev_constant:
            prev_constant = song.constant
            trailing_fragment = 0
            cell_style = 1 if cell_style == 0 else 0
        else:
            trailing_fragment = 0.0000000001 + trailing_fragment
        worksheet.write(
            f"A{index}",
            song.title,
            workbook.add_format(
                {
                    "bg_color": cell_colors[cell_style],
                    "bottom": 5 if idx == len(songs) - 1 else 1,
                    "right": 1,
                    "left": 1,
                }
            ),
        )
        worksheet.write(
            f"B{index}",
            song.difficulty,
            workbook.add_format(
                {
                    "bold": True,
                    "font_color": song.difficulty_color,
                    "bg_color": cell_colors[cell_style],
                    "bottom": 5 if idx == len(songs) - 1 else 1,
                    "right": 1,
                    "left": 1,
                }
            ),
        )
        worksheet.write(
            f"C{index}",
            song.constant + trailing_fragment,
            workbook.add_format(
                {
                    "bg_color": cell_colors[cell_style],
                    "right": 5,
                    "bottom": 5 if idx == len(songs) - 1 else 1,
                    "left": 1,
                    "num_format": "#.0",
                }
            ),
        )
        worksheet.write(
            f"D{index}",
            "",
            workbook.add_format(
                {
                    "num_format": "#'###'##0",
                }
            ),
        )
        worksheet.write(
            f"E{index}",
            f"=MAX(0,IF(D{index}>=10000000,C{index}+2,IF(D{index}>=9800000,C{index}+1+(D{index}-9800000)/200000,C{index}+(D{index}-9500000)/300000)))",
        )

    worksheet = workbook.add_worksheet("RESULTS")
    worksheet.merge_range(
        "A1:E1", "BEST 30 RESULTS", workbook.add_format({"align": "center"})
    )  # type: ignore

    header_style = workbook.add_format(
        {"bold": True, "top": 5, "bottom": 1, "left": 1, "bg_color": "#c9daf8"}
    )
    worksheet.write("A2", "Title", header_style)
    worksheet.set_column("A:A", 30)
    worksheet.write("B2", "DIFF", header_style)
    worksheet.set_column("B:B", 4.35)
    for diff in ["BYD", "ETR", "FTR", "PRS", "PST"]:
        worksheet.conditional_format(
            "B3:B32",
            {
                "type": "cell",
                "criteria": "==",
                "value": f'"{diff}"',
                "format": workbook.add_format(
                    {
                        "bold": True,
                        "font_color": ArcaeaSong("", diff, 0).difficulty_color,
                    }
                ),
            },
        )  # type: ignore
    worksheet.write("C2", "CC", header_style)
    worksheet.set_column("C:C", 4.35)
    worksheet.write("D2", "SCORE", header_style)
    worksheet.set_column("D:D", 14)
    worksheet.write(
        "E2",
        "PLAY RATING",
        workbook.add_format(
            {
                "bold": True,
                "top": 5,
                "bottom": 1,
                "left": 1,
                "right": 5,
                "bg_color": "#c9daf8",
            }
        ),
    )
    worksheet.set_column("E:E", 14)
    for index in range(3, 33):
        worksheet.write(
            f"A{index}",
            f'=IF(E{index}=0," ",INDEX(SCORES!$A:$E,MATCH(E{index},SCORES!$E:$E,0),1))',
            workbook.add_format(
                {
                    "bg_color": "#d9d2e9" if 3 <= index <= 12 else "white",
                    "top": 1,
                    "bottom": 1 if index < 32 else 5,
                    "left": 1,
                    "right": 1,
                }
            ),
        )
        worksheet.write(
            f"B{index}",
            f'=IF(E{index}=0," ",INDEX(SCORES!$A:$E,MATCH(E{index},SCORES!$E:$E,0),2))',
            workbook.add_format(
                {
                    "bg_color": "#d9d2e9" if 3 <= index <= 12 else "white",
                    "top": 1,
                    "bottom": 1 if index < 32 else 5,
                    "left": 1,
                    "right": 1,
                    "bold": True,
                }
            ),
        )
        worksheet.write(
            f"C{index}",
            f'=IF(E{index}=0," ",INDEX(SCORES!$A:$E,MATCH(E{index},SCORES!$E:$E,0),3))',
            workbook.add_format(
                {
                    "bg_color": "#d9d2e9" if 3 <= index <= 12 else "white",
                    "top": 1,
                    "bottom": 1 if index < 32 else 5,
                    "left": 1,
                    "right": 1,
                    "num_format": "#.0",
                }
            ),
        )
        worksheet.write(
            f"D{index}",
            f'=IF(E{index}=0," ",INDEX(SCORES!$A:$E,MATCH(E{index},SCORES!$E:$E,0),4))',
            workbook.add_format(
                {
                    "bg_color": "#d9d2e9" if 3 <= index <= 12 else "white",
                    "top": 1,
                    "bottom": 1 if index < 32 else 5,
                    "left": 1,
                    "right": 1,
                    "num_format": "#'###'##0",
                }
            ),
        )
        worksheet.write(
            f"E{index}",
            f"=LARGE(SCORES!$E:$E,ROW(E{index})-2)",
            workbook.add_format(
                {
                    "bg_color": "#d9d2e9" if 3 <= index <= 12 else "white",
                    "top": 1,
                    "bottom": 1 if index < 32 else 5,
                    "left": 1,
                    "right": 5,
                }
            ),
        )

    worksheet.merge_range(
        "A34:D34",
        "Best 30 average",
        workbook.add_format(
            {
                "top": 5,
                "bottom": 1,
                "right": 1,
                "left": 1,
            }
        ),
    )  # type: ignore
    worksheet.write(
        "E34",
        '=SUMIF(SCORES!$E:$E,">="&LARGE(SCORES!$E:$E,30))/30',
        workbook.add_format(
            {
                "top": 5,
                "bottom": 1,
                "right": 5,
                "left": 1,
            }
        ),
    )
    worksheet.merge_range(
        "A35:D35",
        "Best 10 average (in place of recent top 10 average)",
        workbook.add_format({"top": 1, "bottom": 1, "left": 1, "right": 1}),
    )  # type: ignore
    worksheet.write(
        "E35",
        '=SUMIF(SCORES!$E:$E, ">="&LARGE(SCORES!$E:$E,10))/10',
        workbook.add_format({"top": 1, "bottom": 1, "left": 1, "right": 5}),
    )
    worksheet.merge_range(
        "A36:D36",
        "THEORETICAL MAX PTT",
        workbook.add_format(
            {
                "top": 1,
                "bottom": 5,
                "left": 1,
                "right": 1,
                "bg_color": "yellow",
                "bold": True,
            }
        ),
    )  # type: ignore
    worksheet.write(
        "E36",
        '=IF((E34*30+E35*10)/40>=12.5, CONCATENATE((E34*30+E35*10)/40,"⭐️⭐️"),IF((E34*30+E35*10)/40>=12,CONCATENATE((E34*30+E35*10)/40,"⭐️"),(E34*30+E35*10)/40))',
        workbook.add_format(
            {
                "top": 1,
                "bottom": 5,
                "left": 1,
                "right": 5,
                "bg_color": "yellow",
            }
        ),
    )

    workbook.close()


def main():
    print("Getting level 8-12 song list...")
    songs = scrape_song_list()

    print("Getting level 1-7 song list...")
    songs += scrape_song_list("1-7")

    print("Creating result workbook...")
    construct_workbook(songs, f"Arcaea CC {datetime.now().strftime('%Y-%m-%d')}.xlsx")
    print("Done.")


if __name__ == "__main__":
    exit(main())
