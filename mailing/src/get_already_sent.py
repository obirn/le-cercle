import re


def extract_emails(input_file, output_file):
    with open(input_file, "r") as file:
        emails = re.findall(
            r"\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b", file.read()
        )

    with open(output_file, "w") as file:
        for email in set(emails):
            file.write(email + "\n")


if __name__ == "__main__":
    input_file_path = (
        "mails/noël-2023/sent.txt"  # Change this to the path of your input file
    )
    output_file_path = "output.txt"  # Change this to the desired output file path

    extract_emails(input_file_path, output_file_path)
