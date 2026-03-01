import pandas as pd
import sys
import os

def parse_ppd(ppd_file):
    constraints = []
    options = {}
    option_labels = {}
    current_option = None

    try:
        with open(ppd_file, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()

                if line.startswith("*UIConstraints:"):
                    parts = line.split()
                    pairs = []
                    for i in range(1, len(parts), 2):
                        opt = parts[i].lstrip("*")
                        val = parts[i+1]
                        pairs.append((opt, val))
                    constraints.append(pairs)

                elif line.startswith("*OpenUI"):
                    current_option = line.split()[1].lstrip("*").split("/")[0]
                    options[current_option] = []
                    option_labels[current_option] = {}

                elif current_option and line.startswith(f"*{current_option}"):
                    rest = line.split(maxsplit=1)[1]
                    before_colon = rest.split(":", 1)[0]

                    if "/" in before_colon:
                        internal, label = before_colon.split("/", 1)
                        internal = internal.strip()
                        label = label.strip()
                    else:
                        internal = before_colon.strip()
                        label = internal

                    options[current_option].append(internal)
                    option_labels[current_option][internal] = label

                elif line.startswith("*CloseUI"):
                    current_option = None

    except FileNotFoundError:
        print(f"Error: File '{ppd_file}' not found.")
        sys.exit(1)
    except Exception as e:
        print(f"Error while parsing PPD file: {e}")
        sys.exit(1)

    return constraints, options, option_labels


def build_pivot(option, constraints, options, option_labels):
    if option not in options:
        print(f"Error: Option '{option}' not found in PPD file.")
        print("Available options:", ", ".join(options.keys()))
        sys.exit(1)

    related_opts = set()
    for c in constraints:
        if any(opt == option for opt, val in c):
            related_opts.update([opt for opt, val in c if opt != option])

    rows = []
    for rel_opt in related_opts:
        for rel_val in options.get(rel_opt, []):
            rows.append((rel_opt, rel_val))

    cols = options.get(option, [])

    data = []
    for rel_opt, rel_val in rows:
        row_data = {"RelatedOption": rel_opt, "RelatedValue": rel_val}
        for col_val in cols:
            target_pair = (option, col_val)
            related_pair = (rel_opt, rel_val)
            constrained = any(
                target_pair in c and related_pair in c
                for c in constraints
            )
            row_data[col_val] = "X" if constrained else ""
        data.append(row_data)

    df = pd.DataFrame(data)

    label_row = {"RelatedOption": "", "RelatedValue": ""}
    for col_val in cols:
        label_row[col_val] = option_labels.get(option, {}).get(col_val, "")
    df = pd.concat([pd.DataFrame([label_row]), df], ignore_index=True)

    return df


def print_usage():
    print("Usage: python script.py <ppd_file> <target_option>")
    print("Example: python script.py printer.ppd Duplex")
    sys.exit(1)


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Error: Incorrect number of arguments.")
        print_usage()

    ppd_file = sys.argv[1]
    target_option = sys.argv[2]

    if not os.path.isfile(ppd_file):
        print(f"Error: File '{ppd_file}' does not exist.")
        sys.exit(1)

    constraints, options, option_labels = parse_ppd(ppd_file)
    df = build_pivot(target_option, constraints, options, option_labels)

    print(df.to_string(index=False))

    try:
        # df.to_csv(f"{target_option}_constraints.csv", index=False)
        # print(f"Results saved to '{target_option}_constraints.csv'")

        df.to_excel(f"{target_option}_constraints.xlsx", index=False)
        print(f"Results saved to '{target_option}_constraints.xlsx'")
    except Exception as e:
        print(f"Error saving output files: {e}")
