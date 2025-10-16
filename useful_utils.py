import pendulum
from fuzzywuzzy import fuzz, process
from pandas.io.formats.format import math

import color_utils
import config

month_names_dict = {
    1: "January",
    2: "February",
    3: "March",
    4: "April",
    5: "May",
    6: "June",
    7: "July",
    8: "August",
    9: "September",
    10: "October",
    11: "November",
    12: "December",
}


def get_date_fill(date: int):
    if date % 2 == 0:
        return color_utils.even_header_fill
    else:
        return color_utils.odd_header_fill


def get_date_time_now():
    now = pendulum.now()
    date = now.date()
    date = str(date)
    date = date.split("-")[-1]
    if date.startswith("0"):
        date = date[1:]
    date = int(date)

    time = now.time()
    hour = str(time).split(":")[0]
    hour = int(hour)

    min = str(time).split(":")[1]
    min = int(min)
    time = float(f"{hour}.{min}")
    return (date, time)


def time_min_to_time(time: str) -> str:
    if time.count(".") > 1:
        print(f"{config.failure_string}invalid time provided!")
        return
    if "." not in time:
        return time
    hr, min = time.split(".")
    hr = int(hr)
    min = int(min)
    excess_hr = min // 60
    excess_min = min % 60
    hr = str(hr + excess_hr)
    min = str(excess_min)
    return f"{hr}.{min}"


def get_date_time_next_meal_counting(month_days=31):
    date, time = get_date_time_now()
    next_night_time = float(time_min_to_time(config.count_night_meal_after))
    next_day_time = float(time_min_to_time(config.count_day_meal_after))
    if time >= next_night_time and time < next_day_time:
        return (date, "night")
    else:
        # return night time
        if time <= 24 and time >= 12:
            temp_date = date
            date += 1
            if date > month_days:
                date = date - temp_date
        return (date, "day")


def get_date_color(date):
    if date % 2 == 0:
        return color_utils.even_header_color
    else:
        return color_utils.odd_header_color


# Function to validate a year
def is_valid_year(year):
    try:
        year = int(year)
        if 1900 <= year <= 4000:  # Change the range as needed
            return True
        else:
            return False
    except ValueError:
        return False


def print_in_2_cols(items_list: list, width: int = 30):
    split_point = len(items_list) // 2
    col_1 = items_list[:split_point]
    col_2 = items_list[split_point:]
    for i1, i2 in zip(col_1, col_2):
        print(f"{i1:<{width}}{i2}")
    if len(items_list) % 2 != 0:
        if len(items_list) > 1:
            print(" " * width, end="")
        print(items_list[-1])


def remove_str_from_last(string: str, substr: str):
    string = string.strip()
    len_substr = len(substr)
    if string.endswith(substr):
        string = string[: -1 * len_substr]
    return string


def print_list_upto_width(items_list: list, width: int = 30, seperator: str = ", "):
    items_list = [str(item).strip() for item in items_list]
    lines = []
    line = ""
    for info in items_list:
        if info != items_list[0]:
            line += seperator
        temp_information = line
        line += info
        if len(line) > width:
            temp_information = remove_str_from_last(temp_information, seperator.strip())
            lines.append(temp_information)
            line = info
    if len(line):
        lines.append(remove_str_from_last(line, seperator.strip()))
    for line in lines:
        print(line)


def get_meal_sheet_label_props(labels_list):
    gen_labels = 0
    count_gen = True
    date_labels = 0
    count_date = False
    for header in labels_list:
        if "1_" in header:
            count_gen = False
            count_date = True
        else:
            count_date = False
        if count_gen:
            gen_labels += 1
        if count_date:
            date_labels += 1
        if not count_date and not count_gen:
            break
    date_props_count = date_labels
    gen_prop_count = gen_labels
    days = (len(labels_list[gen_labels:]) // date_labels) + (
        len(labels_list[gen_labels:]) % date_labels
    )

    gen_props = labels_list[1:gen_prop_count]
    label_props = labels_list[gen_prop_count : gen_prop_count + date_props_count]
    date_props = [label.split("_")[-1] for label in label_props]

    return {
        "days": days,
        "date_props_count": date_props_count,
        "gen_props_count": gen_prop_count,
        "gen_props": gen_props,
        "date_props": date_props,
    }


def get_unique_values_from_list(original_list: list):
    unique_list = []
    [unique_list.append(x) for x in original_list if x not in unique_list]
    return unique_list


def final_cofirm(msg: str = "Are you sure?"):
    while True:
        choice = input(f"{config.info_string}{msg} [y/n] ")
        if not choice:
            continue
        if choice in ["y", "Y", "yes", "YES"]:
            return True
        if choice in ["n", "N", "no", "NO"]:
            return False


def take_input(
    question: str = "Input",
    confirm: bool = False,
    cancel_str: str = "00",
):
    ans = None
    while True:
        ans = input(f"{question} [{cancel_str} to cancel] ")
        if not ans:
            continue
        if ans == cancel_str:
            return None
        if confirm:
            confirmed = final_cofirm()
            if confirmed:
                return ans
            else:
                continue
        else:
            return ans


def run_function_continuous(string="want to do it again?", done_str="00"):
    def decorator(func):
        def wrapper(*args, **kwargs):
            done = False
            while True:
                val = func(*args, **kwargs)
                if val is None:
                    break
                while True:
                    print(string)
                    do_continue = input(
                        f"{done_str} for No | anything for yes: "
                    ).strip()
                    if not do_continue:
                        continue
                    if do_continue == done_str:
                        done = True
                    break
                if done:
                    break

        return wrapper

    return decorator


def input_digit_val(string="Enter value(number): ", cancel_str="00"):
    while True:
        inp = input(string).strip()
        if not inp:
            continue
        if inp == cancel_str:
            return
        temp_inp = inp.replace("-", "")
        temp_inp = inp.replace("+", "")
        if temp_inp.isdigit():
            return eval(inp)


def remove_else_add_from_list(original_list: list, item, remove: bool = True):
    unique_list = get_unique_values_from_list(original_list)
    # if present remove it
    if item in unique_list:
        if remove:
            unique_list.remove(item)
    else:
        unique_list.append(item)
    return unique_list


def select_a_item_with_default(
    items_list: list,
    default=None,
    sort: bool = False,
    quit_str: str = "00",
    quit_text: str = "CANCEL",
    input_text: str = None,
    presentation_symbol: str = "*",
    presentation_text: str = "Fuzzy select:",
    width: int = 30,
) -> None | str:
    if sort:
        items_list = sorted(items_list)
    menu_options = []
    if quit_str is not None:
        quit_str = str(quit_str)
        menu_options.append(f"{quit_str}: {quit_text}.")

    def fuzzy_select(options_view_list: list = items_list, info: str = None):
        options_view = options_view_list
        try_counter = 0
        query = None
        while True:
            match = process.extract(
                default or items_list[0], items_list, scorer=fuzz.ratio, limit=2
            )[0]
            choice = match[0]
            if try_counter % 3 == 0:
                print()
                print(presentation_text)
                print(f"{presentation_symbol}" * width * 2)
                print_in_2_cols(options_view, width)
                print()
                print_in_2_cols(menu_options, width)
                print(f"{presentation_symbol}" * width * 2)
            try_counter += 1
            if info:
                print(info)
            input_str = f"{input_text or 'What you want to select?'} (press Enter for `{choice}` or Type) : "
            if query is None:
                query = input(input_str).strip()
                if not query:
                    return choice
            if query == quit_str.strip() and quit_str is not None:
                return
            match = process.extract(query, items_list, scorer=fuzz.ratio, limit=2)[0]
            choice = match[0]
            reselect = input(
                f'is it "{choice}"?\nEnter to confirm. OR re-enter (apprx.) : '
            ).strip()
            if not reselect:
                return choice
            else:
                query = reselect

    return fuzzy_select()


def select_items_from_list(
    items_list: list,
    sort: bool = False,
    multi_select: bool = True,
    confirm_multi_selcet: bool = True,
    quit_str: str = "00",
    quit_text: str = "CANCEL",
    presentation_symbol: str = "*",
    presentation_text: str = "Select:",
    width: int = 30,
    all_select_option: str | None = "0",
    allow_all_selection: bool = True,
    confirm_on_all_selection: bool = False,
    informations_to_show: list | None = None,
    informations_to_show_text: str | None = None,
) -> list[str] | str:
    if sort:
        items_list = sorted(items_list)
    if not multi_select:
        allow_all_selection = False
    if all_select_option == quit_str and allow_all_selection:
        raise Exception(
            f"all selection key, and quit key is same -> {all_select_option}"
        )
    items = []
    menu_options = []
    options_view = [f"{i+1}: {op}" for i, op in enumerate(items_list)]
    if allow_all_selection:
        menu_options.append(f"{all_select_option}: SELECT ALL.")
    if quit_str is not None:
        quit_str = str(quit_str)
        menu_options.append(f"{quit_str}: {quit_text}.")

    choice = None
    try_counter = 0
    while True:
        if try_counter % 3 == 0:
            print()
            print(presentation_text)
            print(f"{presentation_symbol}" * width * 2)
            print_in_2_cols(options_view, width)
            print()
            print_in_2_cols(menu_options, width)
            if informations_to_show is not None:
                seperator = "- " * (width)
                print(seperator)
                if informations_to_show_text is not None:
                    print(informations_to_show_text)
                print_list_upto_width(informations_to_show, width * 2, " | ")
            print(f"{presentation_symbol}" * width * 2)
        try_counter += 1
        input_str = "What you want to do? "
        if multi_select:
            input_str = "Select [eg: 1, 2, 3] "
        choice = input(input_str).strip()
        if choice.strip() == quit_str.strip() and quit_str is not None:
            return
        if not choice:
            continue
        # not multi select
        if not multi_select:
            if not choice.isdigit():
                print(f"{config.info_string}Enter valid option.")
                continue
            choice = int(choice)
            if choice not in list(range(1, len(items_list) + 1)):
                print(f"{config.info_string}Enter valid option.")
                continue
            return items_list[choice - 1]
        # for multi select
        else:
            if choice == all_select_option and allow_all_selection:
                if confirm_on_all_selection:
                    confirmed = final_cofirm()
                    if not confirmed:
                        print(f"{config.info_string}All selection cancelled!")
                        continue
                return items_list
            if choice.endswith(","):
                choice = choice[:-1]
            choices = choice.split(",")
            choices = [ch.strip() for ch in choices]
            wrong_choice = False
            for choice in choices:
                if not choice.isdigit():
                    wrong_choice = True
                    print(f"{choice} is not a valid choice.")
                    break
                if int(choice) not in list(range(1, len(items_list) + 1)):
                    wrong_choice = True
                    print(f"{choice} is not a valid choice.")
                    break
            if wrong_choice:
                continue
            choices = [int(ch.strip()) for ch in choices]
            for choice in choices:
                items.append(items_list[choice - 1])
            if confirm_multi_selcet:
                confirmed = final_cofirm()
                if not confirmed:
                    items = []
                    continue
            return get_unique_values_from_list(items)


def fuzzy_select_from_list(
    items_list: list,
    sort: bool = False,
    multi_select: bool = False,
    confirm_multi_select_text: str = "yy",
    quit_str: str = "00",
    quit_text: str = "CANCEL",
    presentation_symbol: str = "*",
    presentation_text: str = "Fuzzy select:",
    width: int = 30,
):
    if sort:
        items_list = sorted(items_list)
    menu_options = []
    if quit_str is not None:
        quit_str = str(quit_str)
        menu_options.append(f"{quit_str}: {quit_text}.")

    def fuzzy_select(options_view_list: list = items_list, info: str = None):
        options_view = options_view_list
        try_counter = 0
        query = None
        while True:
            if try_counter % 3 == 0:
                print()
                print(presentation_text)
                print(f"{presentation_symbol}" * width * 2)
                print_in_2_cols(options_view, width)
                print()
                print_in_2_cols(menu_options, width)
                print(f"{presentation_symbol}" * width * 2)
            try_counter += 1
            if info:
                print(info)
            input_str = "What you want to select? (approx.) : "
            if query is None:
                query = input(input_str).strip()
                if not query:
                    query = None
                    continue
            if multi_select:
                if (
                    query == confirm_multi_select_text
                    and confirm_multi_select_text in menu_options[-1]
                ):
                    return query
            if query == quit_str.strip() and quit_str is not None:
                return
            match = process.extract(query, items_list, scorer=fuzz.ratio, limit=2)[0]
            choice = match[0]
            reselect = input(
                f'is it "{choice}"?\nEnter to confirm. OR re-enter (apprx.) : '
            ).strip()
            if not reselect:
                return choice
            else:
                query = reselect

    # if multi select
    if multi_select:
        selected = []
        confirm_selection_text = f"{confirm_multi_select_text}: CONFIRM SELECTION."
        info = None
        while True:
            if len(selected) and confirm_selection_text not in menu_options:
                menu_options.append(confirm_selection_text)
            elif not len(selected) and confirm_selection_text in menu_options:
                menu_options = remove_else_add_from_list(
                    menu_options, confirm_selection_text
                )
            result_list = [item for item in items_list if item not in selected]
            choice = fuzzy_select(result_list, info)
            if choice is None:
                return
            if choice == confirm_multi_select_text:
                return selected
            # make unique list
            selected = remove_else_add_from_list(selected, choice)
            if len(selected):
                info = "Selected: "
                info += ", ".join(selected)
                info += f"\n{config.info_string}To unselect, select that item again."
            else:
                info = None
    else:
        return fuzzy_select()


def choose_a_date(days: int, month_string: str, default=30, show_text="date"):
    presentation_text = f"Choose a {show_text} | Month: {month_string.title()}"
    try_counter = 0
    while True:
        if try_counter % 5 == 0:
            print()
            print(presentation_text)
            print("-" * len(presentation_text))
            print(f"    choose between 1 and {days}")
            print()
            print("    00. CANCEL")
            print(f"    defalut is {default}")
        try_counter += 1
        date = input("Enter date: [Blank for default] ").strip()
        if date == "00":
            return
        if not date:
            date = str(default)
        if not date.isdigit():
            print(f"{config.info_string}give correct date.")
            continue
        date = int(date)
        if date < 1 or date > days:
            print(f"{config.info_string}give correct date.")
            continue
        else:
            break
    return date


def choose_time(default="night", show_text="time"):
    times = ["day", "night"]
    if default not in times:
        raise Exception("Provided default is not present in valid times.")
    presentation_text = f"Choose {show_text}"
    try_counter = 0
    while True:
        if try_counter % 5 == 0:
            print()
            print(presentation_text)
            print("-" * len(presentation_text))
            for i, t in enumerate(times):
                print(f"    {i+1}: {t}")
            print()
            print("    00: CANCEL")
            print(f"    defalut is {default}")
        try_counter += 1
        choice = input("Enter time: [Blank for default] ").strip()
        if choice == "00":
            return
        if not choice:
            choice = str(times.index(default) + 1)
        if not choice.isdigit():
            print(f"{config.info_string}give correct number (1/2).")
            continue
        choice = int(choice)
        if choice < 1 or choice > 2:
            print(f"{config.info_string}give correct number (1/2).")
            continue
        else:
            break
    return times[choice - 1]


def get_start_to_end_date_time(date_times: tuple):
    results = []
    if len(date_times) > 2:
        print("err: allowed only 2 items in a tuple -> (item, item)")
        return
    date, time = date_times
    if isinstance(date, int):
        results.append([date, time])
        return results

    froms, tos = date_times
    from_date, from_time = froms
    to_date, to_time = tos
    started = False
    for date in range(from_date, to_date + 1):
        for time in ["day", "night"]:
            if date == from_date and not started:
                if time == from_time:
                    started = True
                    results.append([date, time])
            elif date == to_date:
                results.append([date, time])
                if time == to_time:
                    break
            else:
                results.append([date, time])

    return results


def check_meal_alternate(menu_item, rules):
    """
    Determines whether a meal needs to be altered based on the menu and rules.

    Parameters:
    menu (str): The menu item to check (e.g., "fish", "beef", "chicken").
    rules (str): The rules as a comma-separated string (e.g., "nf,on,od").

    Returns:
    bool: True if the meal needs to be altered, False otherwise.
    """
    # Define abbreviations for restricted items
    restrictions = {
        "fish": config.abbr_no_fish,
        "beef": config.abbr_no_beef,
        "chicken": config.abbr_no_chicken,
        "egg": config.abbr_no_egg,
    }

    # Get the rules as a list (handle None gracefully)
    rule_list = rules.split(",") if rules else []
    rule_list = [r.strip() for r in rule_list]

    # Check if any restriction for the menu item exists in the rules
    if menu_item in restrictions:
        for abbr in restrictions[menu_item]:
            if abbr in rule_list:
                return True

    return False
