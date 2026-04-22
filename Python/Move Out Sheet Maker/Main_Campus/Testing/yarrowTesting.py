from openpyxl import load_workbook, Workbook


def normalize_text(value):
    if value is None:
        return ""
    return str(value).strip()


def get_room_name(name):
    """
    Examples:
    'YARH-1: 101 Bed 1' -> 'YARH-1: 101'
    'YARH-1: 101 Mailbox' -> 'YARH-1: 101'
    """
    parts = normalize_text(name).split()
    if len(parts) >= 2:
        return f"{parts[0]} {parts[1]}"
    return ""


def is_mailbox_row(name):
    return "Mailbox" in normalize_text(name).split()


def is_bed_row(name):
    return "Bed" in normalize_text(name).split()


def load_sheet_as_list(file_path):
    workbook = load_workbook(file_path)
    sheet = workbook.active
    data = []
    for row in sheet.iter_rows(values_only=True):
        data.append(list(row))
    return data


def main():
    key_log_path = input("Enter key log path name: ").strip()
    occupancy_path = input("Enter occupancy path name: ").strip()
    output_file_path = input("Enter the output file path (e.g., output.xlsx): ").strip()

    key_log_list = load_sheet_as_list(key_log_path)
    occupancy_list = load_sheet_as_list(occupancy_path)

    roomWithNoMailKey = {
        "YARH-1: 109",
        "YARH-1: 110",
        "YARH-1: 111",
        "YARH-1: 120",
        "YARH-1: 121",
        "YARH-1: 122",
        "YARH-2: 208",
        "YARH-2: 209",
        "YARH-2: 210",
        "YARH-2: 211",
        "YARH-2: 218",
        "YARH-2: 219",
        "YARH-3: 302",
        "YARH-3: 303",
        "YARH-3: 304",
        "YARH-3: 313",
        "YARH-3: 314",
        "YARH-3: 315",
        "YARH-3: 324",
        "YARH-3: 325",
        "YARH-3: 326",
    }

    # First pass: collect mailbox keys by room
    mailbox_keys = {}
    for line in key_log_list:
        if not line or len(line) < 2:
            continue

        key_name = normalize_text(line[0])
        key_code = line[1]

        if not key_name:
            continue

        if is_mailbox_row(key_name):
            room_name = get_room_name(key_name)
            if room_name:
                mailbox_keys[room_name] = key_code

    # Second pass: build cleaned key rows
    cleaned_rows = []
    missing_mail_keys = []

    for line in key_log_list:
        if not line or len(line) < 2:
            continue

        key_name = normalize_text(line[0])
        key_code = line[1]

        if not key_name:
            continue

        # Skip mailbox source rows in final output
        if is_mailbox_row(key_name):
            continue

        if is_bed_row(key_name):
            room_name = get_room_name(key_name)

            # Add mail key row unless this room is exempt
            if room_name not in roomWithNoMailKey:
                mail_key_code = mailbox_keys.get(room_name, "")
                if mail_key_code == "":
                    missing_mail_keys.append(key_name)

                cleaned_rows.append([
                    f"{key_name} Mail",
                    key_name,
                    "Mail Key",
                    mail_key_code
                ])

            # Add room key row
            cleaned_rows.append([
                f"{key_name} Room",
                key_name,
                "Room Key",
                key_code
            ])
        else:
            # Any non-bed, non-mailbox row
            cleaned_rows.append([
                key_name,
                key_name,
                "Room Key",
                key_code
            ])

    # Build occupancy dictionary: room/bed -> resident name
    occupancy_map = {}
    for row in occupancy_list:
        if not row or len(row) < 2:
            continue

        room_name = normalize_text(row[0])
        resident_name = normalize_text(row[1])

        if room_name:
            occupancy_map[room_name] = resident_name

    # Add resident names to cleaned rows
    residents_name_updated = []
    unmatched_rooms = []

    for row in cleaned_rows:
        full_room_space_description = row[0]
        room_space = normalize_text(row[1])
        key_type_description = row[2]
        key_code = row[3]
        resident_name = occupancy_map.get(room_space, "")

        if resident_name == "":
            unmatched_rooms.append(room_space)

        residents_name_updated.append([
            full_room_space_description,
            room_space,
            key_type_description,
            key_code,
            resident_name
        ])

    # Remove duplicate unmatched entries while keeping order
    seen = set()
    unique_unmatched_rooms = []
    for room in unmatched_rooms:
        if room not in seen:
            seen.add(room)
            unique_unmatched_rooms.append(room)

    # Print results
    for row in residents_name_updated:
        print(row)

    if missing_mail_keys:
        print("\nBeds missing a mailbox key:")
        for item in missing_mail_keys:
            print(item)

    if unique_unmatched_rooms:
        print("\nRooms/Beds with no matching resident found:")
        for item in unique_unmatched_rooms:
            print(item)

    # Save final workbook
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

    for row in residents_name_updated:
        new_sheet.append(row)

    new_workbook.save(output_file_path)
    print(f"\nData successfully saved to {output_file_path}")


if __name__ == "__main__":
    main()