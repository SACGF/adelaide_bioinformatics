#!/usr/bin/env python3

import sys
import pandas as pd


def _to_markdown_row(columns):
    return "| " + " | ".join(columns) + " |"


def main():
    if len(sys.argv) != 2:
        sys.stderr.write(f'Usage: {sys.argv[0]} "Adelaide Bioinformaticians.xls"\n')
        sys.exit(1)
    excel_name = sys.argv[1]

    people_df = pd.read_excel(excel_name, sheet_name="People").fillna("")
    institutions_df = pd.read_excel(excel_name, sheet_name="Institutions", index_col='Code')
    institution_urls = dict(institutions_df[["Institute", "url"]].values)

    # Write header links
    for institute in institutions_df["Institute"].sort_values():
        anchor = institute.lower().replace(" ", "-").replace("/", "")
        print(f"[{institute}](#{anchor})\n")
    print("")

    # Replace with long names
    institution_codes = dict(institutions_df["Institute"])
    people_df["Institute"].replace(institution_codes, inplace=True)

    COLUMNS = ["Name", "email", "ORCID", "student"]
    formatters = {
        "ORCID": lambda orcid: f"[{orcid}](https://orcid.org/{orcid})",
        "student": lambda s: "student" if s else ""
    }

    for institute, people in people_df.groupby("Institute"):
        print(f"# {institute}")
        if url := institution_urls.get(institute):
            print("\n" + url)

        print("")
        bold_columns = [f" **{c}** " for c in COLUMNS]
        header = _to_markdown_row(bold_columns)
        header_spacer = "|-" + "-|-".join([f"-" * len(c) for c in bold_columns]) + "-|"
        print(header)
        print(header_spacer)

        for _, person in people.sort_values("Name").iterrows():
            values = []
            for c in COLUMNS:
                if value := person[c]:
                    if formatter := formatters.get(c):
                        value = formatter(value)
                values.append(value)
            print(_to_markdown_row(values))

        print("")






if __name__ == '__main__':
    main()