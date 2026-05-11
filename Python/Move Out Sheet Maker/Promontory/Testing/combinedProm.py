from openpyxl import load_workbook, Workbook


def load_sheet_as_list(file_path):
    workbook = load_workbook(file_path)
    sheet = workbook.active
    return [list(row) for row in sheet.iter_rows(values_only=True)]


def build_cleaned_key_list(room_list, max_rows=565):
    updated_list = []
    count = 0

    lobby_key_code = " "
    mail_key_code = " "

    for line in room_list:
        if count > max_rows - 1:
            break

        if line and line[0] is not None:
            key_log_name = str(line[0]).split(" ")

            if len(key_log_name) > 2 and key_log_name[2] == "Apt":
                lobby_key_code = line[1]
            elif len(key_log_name) > 2 and key_log_name[2] == "Mailbox":
                mail_key_code = line[1]
            else:
                room_name = str(line[0])
                room_key_code = line[1]

                updated_list.append([room_name + " Apt", room_name, "Apt Key", lobby_key_code])
                updated_list.append([room_name + " Mail", room_name, "Mail Key", mail_key_code])
                updated_list.append([room_name + " Room", room_name, "Room Key", room_key_code])

        count += 1

    return updated_list


def match_residents(occupancy_list, cleaned_list):
    residents_name_updated = []

    for occ_row in occupancy_list:
        if not occ_row or len(occ_row) < 2:
            continue

        room_name = occ_row[0]
        resident_name = occ_row[1]

        if room_name is None:
            continue

        for cleaned_row in cleaned_list:
            if not cleaned_row or len(cleaned_row) < 4:
                continue

            cleaned_room_space = cleaned_row[1]

            if cleaned_room_space is not None and str(room_name) in str(cleaned_room_space):
                residents_name_updated.append([
                    cleaned_row[0],
                    cleaned_row[1],
                    cleaned_row[2],
                    cleaned_row[3],
                    resident_name
                ])

    return residents_name_updated


def save_to_workbook(data, output_file_path):
    new_workbook = Workbook()
    new_sheet = new_workbook.active
    new_sheet.title = "Updated List"

    headers = [
        "Full Room Space Description",
        "Room Space",
        "Key Type Description",
        "Key Code",
        "Residents Name"
    ]
    new_sheet.append(headers)

    for row in data:
        new_sheet.append(row)

    new_workbook.save(output_file_path)


def main():
    key_log_path = input("Enter Key log path name: ")
    occupancy_path = input("Enter Occupancy path Name: ")
    output_file_path = input("Enter the output file path (e.g., output.xlsx): ")

    room_list = load_sheet_as_list(key_log_path)
    occupancy_list = load_sheet_as_list(occupancy_path)

    cleaned_list = build_cleaned_key_list(room_list)
    final_list = match_residents(occupancy_list, cleaned_list)


    save_to_workbook(final_list, output_file_path)
    print(f"Data successfully saved to {output_file_path}")


if __name__ == "__main__":
    main()