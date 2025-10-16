import math
import pendulum
import config
from excel_utils import (
    get_cell_for_boarder_from_df,
    get_val_from_excel_file,
    write_in_excel_file,
)
from useful_utils import (
    choose_a_date,
    choose_time,
    final_cofirm,
    get_meal_sheet_label_props,
    get_start_to_end_date_time,
    input_digit_val,
    month_names_dict,
    run_function_continuous,
    select_items_from_list,
    get_date_time_now,
)
import color_utils as colors


class Boarder:
    def __init__(
        self,
        boarder_values,
        invalid_cells,
        month=None,
        year=None,
        marketing_df=None,
        extra_markeing_df=None,
        file_to_work_on=config.workbook_name,
    ) -> None:
        # self.name = name #name will be parsed from boarder values
        self.boarder_values = boarder_values
        self.invalid_cells = invalid_cells
        self.month = month
        self.year = year
        if self.month is None or self.year is None:
            raise Exception("`Month Number` AND `year` is needed for file editing.")
        self.file_to_work_on = file_to_work_on
        self.boarder_values_headers = list(self.boarder_values.columns)
        self.month_days = get_meal_sheet_label_props(self.boarder_values_headers).get(
            "days"
        )
        self.month_string = month_names_dict[self.month].lower()

        self.marketing_df = marketing_df
        self.extra_markeing_df = extra_markeing_df

        # rules
        self.no_beef = False
        self.no_fish = False
        self.no_egg = False
        self.no_chicken = False
        self.only_night = False
        self.only_day = False
        self.weekend_off = False
        self.weekend_on = False
        self.saturday_night_off = False
        self.sunday_off = False
        self.sunday_on = False
        # status
        self.meal_status = None
        self.guest_meal_status = None
        self.guest_meal_status_count = 0

        # markeing
        self.in_marketing_list = False
        self.total_marketing = 0
        self.total_extra_marketing = 0

        self.parse_values()

    def parse_values(self):
        self.name = self.get_value_from_df("Name")
        self.row_number = (self.boarder_values.index)[0]
        ## meal related
        self.deposit = self.get_deposit()
        self.accessories_charge = self.get_accessories_charge()
        self.total_meal = self.get_total_meal()
        self.total_eggs = self.get_total_eggs()
        self.guest_meals = self.get_guest_meals()
        self.grand_guest_meals = self.get_grand_guest_meals()

        # status
        self.parse_status()
        ## rules
        self.parse_rules()

        # marketing values
        if self.marketing_df is not None:
            self.in_marketing_list = True
            self.total_marketing = self.marketing_df["Money Spent"].sum()
        if self.extra_markeing_df is not None:
            self.total_extra_marketing = self.extra_markeing_df["Money Spent"].sum()

    def menu(self, quit_str="00"):
        options_table = [
            ["Deposit", self.do_deposit],
            ["turn on meal", self.change_meal_on],
            ["turn off meal", self.change_meal_off],
            ["turn on guest meal", self.change_guest_meal_on],
            ["turn off guest meal", self.change_guest_meal_off],
            ["put/change extra egg", self.change_extra_egg_count],
            ["change meal status(auto calculation)", self.change_status],
            ["change rules(eg: od, on etc)", self.change_rules],
            ["Accessories Charge(laptop, fan etc)", self.work_with_accessories_charge],
            ["reset cell values[DANGER]", self.reset_cell_values],
        ]

        options = [op[0] for op in options_table]
        informations_to_show = self.get_informations_to_show()
        choice = None
        date_now, _ = get_date_time_now()
        while True:
            choice = select_items_from_list(
                options,
                multi_select=False,
                quit_str=quit_str,
                quit_text="DONE",
                presentation_text=f"Menu for: {self.name} | {self.month_string.title()}-{date_now}",
                width=40,
                informations_to_show=informations_to_show,
            )
            if choice is None:
                break
            choice_idx = options.index(choice)
            options_table[choice_idx][1]()
            self.parse_values()
            informations_to_show = self.get_informations_to_show()
            date_now, _ = get_date_time_now()

    def get_informations_to_show(self):
        infos = []
        if not self.in_marketing_list:
            infos.append(f"{config.warning_string}Not in marketing list.")
        meal_status_txt = "on" if self.meal_status else "off"
        infos.append(f"meal status: {meal_status_txt}")
        infos.append(f"deposit: {config.rupee_symbol}{self.deposit}")
        meal_today = self.get_total_meal_today()
        if meal_today is not None:
            if meal_today != self.total_meal:
                infos.append(f"meal till now: {meal_today}")
        infos.append(f"total meal: {self.total_meal}")
        if self.total_eggs:
            infos.append(f"total eggs: {self.total_eggs}")
        if self.guest_meals:
            infos.append(f"total guests: {self.guest_meals}")
        if self.grand_guest_meals:
            infos.append(f"total grand guests: {self.grand_guest_meals}")
        if self.accessories_charge:
            infos.append(f"Acs Charge: {config.rupee_symbol}{self.accessories_charge}")
        if self.total_marketing:
            infos.append(
                f"Markeing amount: {config.rupee_symbol}{self.total_marketing}"
            )
        if self.total_extra_marketing:
            infos.append(
                f"Extra-markeing amount: {config.rupee_symbol}{self.total_extra_marketing}"
            )

        rules_str = []
        if self.no_beef:
            rules_str.append("no beef")
        if self.no_fish:
            rules_str.append("no-fish")
        if self.no_egg:
            rules_str.append("no-egg")
        if self.no_chicken:
            rules_str.append("no-chicken")
        if self.only_night:
            rules_str.append("only-night")
        if self.only_day:
            rules_str.append("only-day")
        if self.weekend_off:
            rules_str.append("weekend-off")
        if self.weekend_on:
            rules_str.append("weekend-on")
        if self.saturday_night_off:
            rules_str.append("sat-night-off")
        if self.sunday_off:
            rules_str.append("sunday-off")
        if self.sunday_on:
            rules_str.append("sunday-on")
        if len(rules_str):
            infos.append(f"{', '.join(rules_str)}")

        if self.guest_meal_status:
            infos.append(f"auto-guest-meal({self.guest_meal_status_count})")

        return infos

    def change_meal_on(self):
        date_time = self.select_date_and_time(
            show_text=f"date for Meal on | {self.name}"
        )
        if date_time is None:
            return
        date_times = get_start_to_end_date_time(date_time)
        for date, time in date_times:
            self.turn_on_meal(
                date, time, write_eggs=False, write_guest_meal=False, overwrite=True
            )

    def change_meal_off(self):
        date_time = self.select_date_and_time(
            show_text=f"date for Meal off | {self.name}"
        )
        if date_time is None:
            return
        date_times = get_start_to_end_date_time(date_time)
        for date, time in date_times:
            self.turn_off_meal(date, time, overwrite=True)

    def change_guest_meal_on(self):
        date_time = self.select_date_and_time(
            show_text=f"date for Guest Meal on | {self.name}"
        )
        if date_time is None:
            return
        date_times = get_start_to_end_date_time(date_time)
        for date, time in date_times:
            amount = input_digit_val(f"Enter guest meal count({date}|{time}): ")
            self.turn_on_guest_meal(date, time, amount=amount, overwrite=True)

    def change_guest_meal_off(self):
        date_time = self.select_date_and_time(
            show_text=f"date for Guest Meal off | {self.name}"
        )
        if date_time is None:
            return
        date_times = get_start_to_end_date_time(date_time)
        for date, time in date_times:
            self.turn_off_guest_meal(date, time, overwrite=True)

    def change_extra_egg_count(self):
        date_time = self.select_date_and_time(
            show_text=f"date for put/change extra eggs | {self.name}"
        )
        if date_time is None:
            return
        date_times = get_start_to_end_date_time(date_time)
        for date, time in date_times:
            amount = input_digit_val(f"Enter egg count({date}|{time}): ")
            self.change_extra_egg(date, time, value=amount)

    def parse_status(self):
        header = "Status"
        cell = get_cell_for_boarder_from_df(self.boarder_values, self.name, header)
        if cell in self.invalid_cells:
            return
        status = self.get_value_from_df("Status")
        if not status:
            self.meal_status = True

        self.guest_meal_status = False
        status = status.split(",")
        status = [val.strip() for val in status]
        for stat in status:
            if stat in config.abbr_status_on:
                self.meal_status = True
            elif stat in config.abbr_status_off:
                self.meal_status = False
            elif stat in config.abbr_continue_guest_meal:
                self.guest_meal_status = True
                for st in status:
                    if st.isdigit():
                        self.guest_meal_status_count = int(st)

    def parse_rules(self):
        rules = self.get_value_from_df("Rules")
        if not rules:
            return
        rules = rules.split(",")
        rules = [r.strip().lower() for r in rules]
        for rule in rules:
            if rule in config.abbr_no_beef:
                self.no_beef = True
            if rule in config.abbr_no_fish:
                self.no_fish = True
            if rule in config.abbr_no_egg:
                self.no_egg = True
            if rule in config.abbr_no_chicken:
                self.no_chicken = True
            if rule in config.abbr_only_night:
                self.only_night = True
            if rule in config.abbr_only_day:
                self.only_day = True
            if rule in config.abbr_weekend_off:
                self.weekend_off = True
            if rule in config.abbr_weekend_on:
                self.weekend_on = True
            if rule in config.abbr_saturday_night_off:
                self.saturday_night_off = True
            if rule in config.abbr_sunday_off:
                self.sunday_off = True
            if rule in config.abbr_sunday_on:
                self.sunday_on = True

    ############## Editing
    def turn_off_meal(self, date: int, time: str, overwrite: bool = False):
        allowed_times = ["day", "night"]
        if time not in allowed_times:
            print(f"{config.failure_string}provided time is invalid - {time}")
            return
        if (
            self.meal_status is None
            or self.guest_meal_status is None
            or self.total_meal is None
            or self.guest_meals is None
            or self.total_eggs is None
        ):
            print(
                f"{config.failure_string}invalid value is present for `{self.name}`. correct it."
            )
            print(f"{config.failure_string}can not turn off meal")
            return
        off_val = config.abbr_meal_off[0]
        df_label = self.make_df_label(date, time)
        if not overwrite:
            val = self.get_value_from_df(df_label)
            if val:
                print(
                    f"{config.info_string}value already present for {self.name} - {df_label}. skipping."
                )
                return False
        current_val = self.get_value_from_df(df_label)
        set_df_value = False
        if not current_val:
            set_df_value = True
        elif current_val in config.abbr_meal_on:
            set_df_value = True

        cell = get_cell_for_boarder_from_df(self.boarder_values, self.name, df_label)
        write_in_excel_file(
            config.meal_count_sheet_name,
            cell,
            off_val,
            center_aligned=True,
            color_reset=True,
            fill_reset=True,
        )
        print(
            f"{config.meal_off_string}turned off meal for {self.name}. date- {date}|{time}."
        )
        if set_df_value:
            self.set_value_to_df(df_label, off_val)
            self.total_meal -= 1
        return True

    def turn_on_meal(
        self,
        date: int,
        time: str,
        follow_rule: bool = False,
        overwrite: bool = False,
        write_guest_meal: bool = True,
        write_eggs: bool = True,
    ):
        print()
        allowed_times = ["day", "night"]
        date_today = int(pendulum.today().format("DD"))
        date_str_show = date_today
        if date_today == date:
            date_str_show = "Today"
        if date == date_today - 1:
            date_str_show = "Yesterday"
        if date == date_today + 1:
            date_str_show = "Tomorrow"

        if time not in allowed_times:
            print(f"{config.failure_string}provided time is invalid - {time}")
            return
        if (
            self.meal_status is None
            or self.guest_meal_status is None
            or self.total_meal is None
            or self.guest_meals is None
            or self.total_eggs is None
        ):
            print(
                f"{config.failure_string}invalid value is present for `{self.name}`. correct it."
            )
            print(f"{config.failure_string}can not add meal")
            return
        df_label = self.make_df_label(date, time)
        if not overwrite:
            val = self.get_value_from_df(df_label)
            if (
                self.meal_status is False
                and val in config.abbr_meal_on
                and date >= date_today
            ):
                print(
                    f"{config.info_string}{self.name}|meal status:off, but meal is on for {date_str_show}|{time}"
                )
                confirmed = final_cofirm(
                    f"Do you want to change meal-status to continue-on for {self.name}?"
                )
                if confirmed:
                    self.meal_status = True
                    return self.change_meal_status()
                else:
                    print(
                        f"{config.info_string}value already present for {self.name} - {df_label}. skipping."
                    )
                    return False
            else:
                if val:
                    print(
                        f"{config.info_string}value already present for {self.name} - {df_label}. skipping."
                    )
                    return False
        current_val = self.get_value_from_df(df_label)
        set_df_value = False
        if not current_val:
            set_df_value = True
        elif current_val in config.abbr_meal_off:
            set_df_value = True
        on_val = config.abbr_meal_on[0]
        cell = get_cell_for_boarder_from_df(self.boarder_values, self.name, df_label)
        # general meal
        if follow_rule:
            date_str = date
            if len(str(date_str)) == 1:
                date_str = f"0{date_str}"
            month_str = self.month
            if len(str(month_str)) == 1:
                month_str = f"0{month_str}"
            date_string = f"{self.year}-{month_str}-{date_str}"
            day_string = pendulum.parse(date_string).format("dddd").lower()
            turn_off_meal = False
            # day
            if time == allowed_times[0]:
                if self.only_night:
                    turn_off_meal = True
            # night
            if time == allowed_times[1]:
                if self.only_day:
                    turn_off_meal = True
            ## other rules
            # saturday
            if day_string == "saturday":
                if self.weekend_off:
                    turn_off_meal = True
                if self.weekend_on:
                    turn_off_meal = False
                # saturday night
                if time == allowed_times[1]:
                    if self.saturday_night_off:
                        turn_off_meal = True
            # sunday
            if day_string == "sunday":
                if self.weekend_off:
                    turn_off_meal = True
                if self.weekend_on:
                    turn_off_meal = False
                if self.sunday_on:
                    turn_off_meal = False

            # meal status
            if self.meal_status is False:
                turn_off_meal = True

            if turn_off_meal:
                self.turn_off_meal(date, time, overwrite=overwrite)
            else:
                # finally
                write_in_excel_file(
                    config.meal_count_sheet_name,
                    cell,
                    on_val,
                    center_aligned=True,
                    color=colors.meal_on_color,
                    fill=colors.meal_on_fill,
                )
                print(
                    f"{config.meal_on_string}turned on meal for {self.name}. date- {date}|{time}."
                )

        else:
            write_in_excel_file(
                config.meal_count_sheet_name,
                cell,
                on_val,
                center_aligned=True,
                color=colors.meal_on_color,
                fill=colors.meal_on_fill,
            )
            print(
                f"{config.meal_on_string}turned on meal for {self.name}. date- {date}|{time}."
            )
        if set_df_value:
            self.set_value_to_df(df_label, on_val)
            self.total_meal += 1
        # guest meal
        if write_guest_meal:
            self.turn_on_guest_meal(date, time, overwrite=overwrite)
        # put egg count to 0
        if write_eggs:
            self.change_extra_egg(date, time, value=0, notify_on_zero=False)

        print()
        return True

    def change_extra_egg(self, date: int, time: str, value=None, notify_on_zero=True):
        if self.total_eggs is None:
            print(f"{config.failure_string}eggs value is inavalid. Correct it.")
            return
        # get current egg of the day
        final_val = value
        df_label = self.make_df_label(date, f"egg({time})")
        if value is None:
            val = self.get_value_from_df(df_label)
            final_val = val
            if not val:
                final_val = 0
            options = ["add to current egg count", "change current egg count"]
            choice = select_items_from_list(
                options,
                multi_select=False,
                presentation_text=f"Egg count of {self.name} of date- ({date},{time}) | Current eggs: {final_val}",
            )
            if choice is None:
                return
            if choice == options[0]:
                extra_eggs_value = input_digit_val(
                    f"Enter extra egg amount of {self.name} for day- {date}|{time}: "
                )
                if extra_eggs_value is None:
                    return
                final_val += extra_eggs_value
            if choice == options[1]:
                new_eggs_value = input_digit_val(
                    f"Enter total egg amount of {self.name} for day- {date}|{time}:  "
                )
                if new_eggs_value is None:
                    return
                final_val = new_eggs_value
        cell = get_cell_for_boarder_from_df(self.boarder_values, self.name, df_label)
        if final_val == 0:
            write_in_excel_file(
                config.meal_count_sheet_name,
                cell,
                final_val,
                center_aligned=True,
                color_reset=True,
                fill_reset=True,
            )

            if notify_on_zero:
                print(
                    f"{config.success_string}updated total egg count to {final_val} for {self.name}. date- {date}|{time}."
                )
        else:
            write_in_excel_file(
                config.meal_count_sheet_name,
                cell,
                final_val,
                center_aligned=True,
                color=colors.egg_amount_color,
                fill=colors.egg_amount_fill,
            )
            print(
                f"{config.success_string}updated total egg count to {final_val} for {self.name}. date- {date}|{time}."
            )
        self.set_value_to_df(df_label, final_val)
        return True

    def turn_off_guest_meal(self, date: int, time: str, overwrite: bool = False):
        if (
            self.meal_status is None
            or self.guest_meal_status is None
            or self.guest_meals is None
        ):
            print(
                f"{config.failure_string}invalid value is present for `{self.name}`. correct it."
            )
            print(f"{config.failure_string}can not add guest meal")
            return
        # add guest meal
        final_amount = 0
        df_label = self.make_df_label(date, f"guest({time})")
        if df_label not in self.boarder_values_headers:
            df_label = self.make_df_label(date, f"grand_guest({time})")
        if not overwrite:
            val = self.get_value_from_df(df_label)
            if val:
                print(
                    f"{config.info_string}value already present for {self.name} - {df_label}. skipping."
                )
                return False
        cell = get_cell_for_boarder_from_df(self.boarder_values, self.name, df_label)
        write_in_excel_file(
            config.meal_count_sheet_name,
            cell,
            final_amount,
            center_aligned=True,
            color_reset=True,
            fill_reset=True,
        )
        print(
            f"{config.meal_off_string}turned off guest meal for {self.name}. date- {date}|{time}."
        )
        return True

    def turn_on_guest_meal(
        self, date: int, time: str, amount: int = None, overwrite: bool = False
    ):
        if (
            self.meal_status is None
            or self.guest_meal_status is None
            or self.guest_meals is None
        ):
            print(
                f"{config.failure_string}invalid value is present for `{self.name}`. correct it."
            )
            print(f"{config.failure_string}can not add guest meal")
            return
        final_amount = amount
        if not amount:
            final_amount = self.guest_meal_status_count
            if final_amount == 0:
                return self.turn_off_guest_meal(date, time, overwrite)
        # add guest meal
        df_label = self.make_df_label(date, f"guest({time})")
        if df_label not in self.boarder_values_headers:
            df_label = self.make_df_label(date, f"grand_guest({time})")
        if not overwrite:
            val = self.get_value_from_df(df_label)
            if val:
                print(
                    f"{config.info_string}value already present for {self.name} - {df_label}. skipping."
                )
                return False
        cell = get_cell_for_boarder_from_df(self.boarder_values, self.name, df_label)
        write_in_excel_file(
            config.meal_count_sheet_name,
            cell,
            final_amount,
            center_aligned=True,
            color=colors.guest_amount_color,
            fill=colors.guest_amount_fill,
        )
        self.set_value_to_df(df_label, final_amount)
        print(
            f"{config.meal_on_string}turned on guest meal for {self.name}. date- {date}|{time}. amount- {final_amount}"
        )
        return True

    @run_function_continuous("continue work with deposit?")
    def do_deposit(self):
        if self.deposit is None:
            print(f"{config.failure_string}Deposit value is inavalid. Correct it.")
            return
        final_deposit = self.deposit
        options = ["add to deposit", "change deposit"]
        choice = select_items_from_list(
            options,
            multi_select=False,
            presentation_text=f"Deposit for {self.name} | Current deposit: {config.rupee_symbol}{self.deposit}",
        )
        if choice is None:
            return
        if choice == options[0]:
            extra_deposit = input_digit_val(
                f"Enter additional dposit amount for {self.name}: {config.rupee_symbol}"
            )
            if extra_deposit is None:
                return
            final_deposit += extra_deposit
        if choice == options[1]:
            new_deposit = input_digit_val(
                f"Enter total Deposit of {self.name}: {config.rupee_symbol}"
            )
            if new_deposit is None:
                return
            final_deposit = new_deposit

        self.deposit = final_deposit
        cell = get_cell_for_boarder_from_df(self.boarder_values, self.name, "Deposit")
        write_in_excel_file(config.meal_count_sheet_name, cell, final_deposit)
        self.set_value_to_df("Deposit", final_deposit)
        print(
            f"{config.success_string}changed deposit for {self.name}. amount- {config.rupee_symbol}{final_deposit}"
        )
        return True

    @run_function_continuous("continue work with Accessories Charge?")
    def work_with_accessories_charge(self):
        if self.accessories_charge is None:
            print(f"{config.failure_string}Acs Charge value is inavalid. Correct it.")
            return
        final_acs_charge = self.accessories_charge
        options = ["add to accessories charge", "change accessories charge"]
        choice = select_items_from_list(
            options,
            multi_select=False,
            presentation_text=f"Accessories Charge of {self.name} | Current charge: {self.accessories_charge}",
        )
        if choice is None:
            return
        if choice == options[0]:
            extra_charge = input_digit_val(f"Enter additional amount for {self.name}: ")
            if extra_charge is None:
                return
            final_acs_charge += extra_charge
        if choice == options[1]:
            new_charge = input_digit_val(
                f"Enter total Accessories Charge of {self.name}: "
            )
            if new_charge is None:
                return
            final_acs_charge = new_charge

        header = "Acs Charge"
        self.accessories_charge = final_acs_charge
        cell = get_cell_for_boarder_from_df(self.boarder_values, self.name, header)
        write_in_excel_file(config.meal_count_sheet_name, cell, final_acs_charge)
        self.set_value_to_df(header, final_acs_charge)
        print(
            f"{config.success_string}changed accessories charge for {self.name}. amount- {final_acs_charge}"
        )
        return True

    def change_rules(self):
        no_beef_text = "off" if self.no_beef else "on"
        no_fish_text = "off" if self.no_fish else "on"
        no_egg_text = "off" if self.no_egg else "on"
        no_chicken_text = "off" if self.no_chicken else "on"
        only_night_text = "off" if self.only_night else "on"
        only_day_text = "off" if self.only_day else "on"
        weekend_off_text = "off" if self.weekend_off else "on"
        weekend_on_text = "off" if self.weekend_on else "on"
        saturday_night_off_text = "off" if self.saturday_night_off else "on"
        sunday_off_text = "off" if self.sunday_off else "on"
        sunday_on_text = "off" if self.sunday_on else "on"

        options_list = [
            [f"turn {no_beef_text} no-beef", self.toggle_no_beef],
            [f"turn {no_fish_text} no-fish", self.toggle_no_fish],
            [f"turn {no_egg_text} no-egg", self.toggle_no_egg],
            [f"turn {no_chicken_text} no-chicken", self.toggle_no_chicken],
            [f"turn {only_night_text} only-night", self.toggle_only_night],
            [f"turn {only_day_text} only-day", self.toggle_only_day],
            [f"turn {weekend_on_text} weekend-on", self.toggle_weekend_on],
            [f"turn {weekend_off_text} weekend-off", self.toggle_weekend_off],
            [
                f"turn {saturday_night_off_text} saturday-night-off",
                self.toggle_saturday_night_off,
            ],
            [f"turn {sunday_off_text} sunday-off", self.toggle_sunday_off],
            [f"turn {sunday_on_text} sunday-on", self.toggle_sunday_on],
        ]
        options = [op[0] for op in options_list]
        infos_on = []
        for op in options:
            if " off " in op:
                infos_on.append(op.split(" ")[-1])
        if not len(infos_on):
            infos_on = None

        choice = select_items_from_list(
            options,
            multi_select=True,
            quit_text="EXIT",
            presentation_text=f"Rules for {self.name}",
            informations_to_show=infos_on,
            informations_to_show_text="TURNED ON RULES:",
            allow_all_selection=False,
        )
        if choice is None:
            return
        for ch in choice:
            choice_idx = options.index(ch)
            options_list[choice_idx][1]()

    def toggle_no_beef(self):
        header = "Rules"
        cell = get_cell_for_boarder_from_df(self.boarder_values, self.name, header)
        rules = get_val_from_excel_file(config.meal_count_sheet_name, cell)
        final_rules = []
        if rules:
            rules = [rule.strip() for rule in rules.split(",")]
            for rule in rules:
                if rule not in config.abbr_no_beef:
                    final_rules.append(rule)
        on_val = config.abbr_no_beef[0]
        stat_text = ""
        if self.no_beef:
            # turn off
            stat_text = "removed"
            self.no_beef = False
        else:
            # turn on
            stat_text = "applied"
            final_rules.append(on_val)
            self.no_beef = True
        final_rules = ", ".join(final_rules)
        write_in_excel_file(config.meal_count_sheet_name, cell, final_rules)
        print(f"{config.success_string}no-beef {stat_text} for {self.name}.")
        return True

    def toggle_no_fish(self):
        header = "Rules"
        cell = get_cell_for_boarder_from_df(self.boarder_values, self.name, header)
        rules = get_val_from_excel_file(config.meal_count_sheet_name, cell)
        final_rules = []
        if rules:
            rules = [rule.strip() for rule in rules.split(",")]
            for rule in rules:
                if rule not in config.abbr_no_fish:
                    final_rules.append(rule)
        on_val = config.abbr_no_fish[0]
        stat_text = ""
        if self.no_fish:
            # turn off
            stat_text = "removed"
            self.no_fish = False
        else:
            # turn on
            stat_text = "applied"
            final_rules.append(on_val)
            self.no_fish = True
        final_rules = ", ".join(final_rules)
        write_in_excel_file(config.meal_count_sheet_name, cell, final_rules)
        print(f"{config.success_string}no-fish {stat_text} for {self.name}.")
        return True

    def toggle_no_egg(self):
        header = "Rules"
        cell = get_cell_for_boarder_from_df(self.boarder_values, self.name, header)
        rules = get_val_from_excel_file(config.meal_count_sheet_name, cell)
        final_rules = []
        if rules:
            rules = [rule.strip() for rule in rules.split(",")]
            for rule in rules:
                if rule not in config.abbr_no_egg:
                    final_rules.append(rule)
        on_val = config.abbr_no_egg[0]
        stat_text = ""
        if self.no_egg:
            # turn off
            stat_text = "removed"
            self.no_egg = False
        else:
            # turn on
            stat_text = "applied"
            final_rules.append(on_val)
            self.no_egg = True
        final_rules = ", ".join(final_rules)
        write_in_excel_file(config.meal_count_sheet_name, cell, final_rules)
        print(f"{config.success_string}no-egg {stat_text} for {self.name}.")
        return True

    def toggle_no_chicken(self):
        header = "Rules"
        cell = get_cell_for_boarder_from_df(self.boarder_values, self.name, header)
        rules = get_val_from_excel_file(config.meal_count_sheet_name, cell)
        final_rules = []
        if rules:
            rules = [rule.strip() for rule in rules.split(",")]
            for rule in rules:
                if rule not in config.abbr_no_chicken:
                    final_rules.append(rule)
        on_val = config.abbr_no_chicken[0]
        stat_text = ""
        if self.no_chicken:
            # turn off
            stat_text = "removed"
            self.no_chicken = False
        else:
            # turn on
            stat_text = "applied"
            final_rules.append(on_val)
            self.no_chicken = True
        final_rules = ", ".join(final_rules)
        write_in_excel_file(config.meal_count_sheet_name, cell, final_rules)
        print(f"{config.success_string}no-chicken {stat_text} for {self.name}.")
        return True

    def toggle_only_night(self):
        header = "Rules"
        cell = get_cell_for_boarder_from_df(self.boarder_values, self.name, header)
        rules = get_val_from_excel_file(config.meal_count_sheet_name, cell)
        final_rules = []
        if rules:
            rules = [rule.strip() for rule in rules.split(",")]
            for rule in rules:
                if rule not in config.abbr_only_night:
                    final_rules.append(rule)
        on_val = config.abbr_only_night[0]
        stat_text = ""
        if self.only_night:
            # turn off
            stat_text = "removed"
            self.only_night = False
        else:
            # turn on
            stat_text = "applied"
            final_rules.append(on_val)
            self.only_night = True
        final_rules = ", ".join(final_rules)
        write_in_excel_file(config.meal_count_sheet_name, cell, final_rules)
        print(f"{config.success_string}only-night {stat_text} for {self.name}.")
        return True

    def toggle_only_day(self):
        header = "Rules"
        cell = get_cell_for_boarder_from_df(self.boarder_values, self.name, header)
        rules = get_val_from_excel_file(config.meal_count_sheet_name, cell)
        final_rules = []
        if rules:
            rules = [rule.strip() for rule in rules.split(",")]
            for rule in rules:
                if rule not in config.abbr_only_day:
                    final_rules.append(rule)
        on_val = config.abbr_only_day[0]
        stat_text = ""
        if self.only_day:
            # turn off
            stat_text = "removed"
            self.only_day = False
        else:
            # turn on
            stat_text = "applied"
            final_rules.append(on_val)
            self.only_day = True
        final_rules = ", ".join(final_rules)
        write_in_excel_file(config.meal_count_sheet_name, cell, final_rules)
        print(f"{config.success_string}only-day {stat_text} for {self.name}.")
        return True

    def toggle_weekend_on(self):
        header = "Rules"
        cell = get_cell_for_boarder_from_df(self.boarder_values, self.name, header)
        rules = get_val_from_excel_file(config.meal_count_sheet_name, cell)
        final_rules = []
        if rules:
            rules = [rule.strip() for rule in rules.split(",")]
            for rule in rules:
                if rule not in config.abbr_weekend_on:
                    final_rules.append(rule)
        on_val = config.abbr_weekend_on[0]
        stat_text = ""
        if self.weekend_on:
            # turn off
            stat_text = "removed"
            self.weekend_on = False
        else:
            # turn on
            stat_text = "applied"
            final_rules.append(on_val)
            self.weekend_on = True
        final_rules = ", ".join(final_rules)
        write_in_excel_file(config.meal_count_sheet_name, cell, final_rules)
        print(f"{config.success_string}weekend-on {stat_text} for {self.name}.")
        return True

    def toggle_weekend_off(self):
        header = "Rules"
        cell = get_cell_for_boarder_from_df(self.boarder_values, self.name, header)
        rules = get_val_from_excel_file(config.meal_count_sheet_name, cell)
        final_rules = []
        if rules:
            rules = [rule.strip() for rule in rules.split(",")]
            for rule in rules:
                if rule not in config.abbr_weekend_off:
                    final_rules.append(rule)
        on_val = config.abbr_weekend_off[0]
        stat_text = ""
        if self.weekend_off:
            # turn off
            stat_text = "removed"
            self.weekend_off = False
        else:
            # turn on
            stat_text = "applied"
            final_rules.append(on_val)
            self.weekend_off = True
        final_rules = ", ".join(final_rules)
        write_in_excel_file(config.meal_count_sheet_name, cell, final_rules)
        print(f"{config.success_string}weekend-off {stat_text} for {self.name}.")
        return True

    def toggle_saturday_night_off(self):
        header = "Rules"
        cell = get_cell_for_boarder_from_df(self.boarder_values, self.name, header)
        rules = get_val_from_excel_file(config.meal_count_sheet_name, cell)
        final_rules = []
        if rules:
            rules = [rule.strip() for rule in rules.split(",")]
            for rule in rules:
                if rule not in config.abbr_saturday_night_off:
                    final_rules.append(rule)
        on_val = config.abbr_saturday_night_off[0]
        stat_text = ""
        if self.saturday_night_off:
            # turn off
            stat_text = "removed"
            self.saturday_night_off = False
        else:
            # turn on
            stat_text = "applied"
            final_rules.append(on_val)
            self.saturday_night_off = True
        final_rules = ", ".join(final_rules)
        write_in_excel_file(config.meal_count_sheet_name, cell, final_rules)
        print(f"{config.success_string}saturday-night-off {stat_text} for {self.name}.")
        return True

    def toggle_sunday_off(self):
        header = "Rules"
        cell = get_cell_for_boarder_from_df(self.boarder_values, self.name, header)
        rules = get_val_from_excel_file(config.meal_count_sheet_name, cell)
        final_rules = []
        if rules:
            rules = [rule.strip() for rule in rules.split(",")]
            for rule in rules:
                if rule not in config.abbr_sunday_off:
                    final_rules.append(rule)
        on_val = config.abbr_sunday_off[0]
        stat_text = ""
        if self.sunday_off:
            # turn off
            stat_text = "removed"
            self.sunday_off = False
        else:
            # turn on
            stat_text = "applied"
            final_rules.append(on_val)
            self.sunday_off = True
        final_rules = ", ".join(final_rules)
        write_in_excel_file(config.meal_count_sheet_name, cell, final_rules)
        print(f"{config.success_string}sunday-off {stat_text} for {self.name}.")
        return True

    def toggle_sunday_on(self):
        header = "Rules"
        cell = get_cell_for_boarder_from_df(self.boarder_values, self.name, header)
        rules = get_val_from_excel_file(config.meal_count_sheet_name, cell)
        final_rules = []
        if rules:
            rules = [rule.strip() for rule in rules.split(",")]
            for rule in rules:
                if rule not in config.abbr_sunday_on:
                    final_rules.append(rule)
        on_val = config.abbr_sunday_on[0]
        stat_text = ""
        if self.sunday_on:
            # turn off
            stat_text = "removed"
            self.sunday_on = False
        else:
            # turn on
            stat_text = "applied"
            final_rules.append(on_val)
            self.sunday_on = True
        final_rules = ", ".join(final_rules)
        write_in_excel_file(config.meal_count_sheet_name, cell, final_rules)
        print(f"{config.success_string}sunday-on {stat_text} for {self.name}.")
        return True

    def reset_cell_values(self):
        options_meal = [
            ["meal: select date", self.reset_cell_meal_date],
            ["meal: all off values", self.reset_cell_meal_all],
        ]
        options_guest_meal = [
            ["guest-meal: select date", self.reset_cell_guest_meal_date],
            ["guest-meal: all off", self.reset_cell_guest_meal_all],
        ]
        options_table = options_meal + options_guest_meal
        options = [op[0] for op in options_table]
        print(
            f"{config.warning_string}YOU WILL RESET CELLS WITH OFF VALUE. PROCEED WITH CAUTION."
        )
        input("press any key")
        choice = select_items_from_list(
            options,
            multi_select=False,
            presentation_text=f"RESET CELLS (only off values) {config.warning_string}be sure before performing any actions.",
        )
        if choice is None:
            return
        choice_idx = options.index(choice)
        options_table[choice_idx][1]()

    def reset_cell_meal_date(self):
        date_time = self.select_date_and_time(show_text=f"reset date for {self.name}")
        if date_time is None:
            return
        date_times = get_start_to_end_date_time(date_time)
        df_headers = []
        for date, time in date_times:
            header_df = f"{date}_{time}"
            val = self.get_value_from_df(header_df)
            if val in config.abbr_meal_off:
                df_headers.append(header_df)
        print(
            f"{config.warning_string} you are about to reset {len(df_headers)} meal cells for {self.name}"
        )
        confirmed = final_cofirm()
        if not confirmed:
            return
        self.reset_cell(df_headers)

    def reset_cell_meal_all(self):
        date_time = ((1, "day"), (self.month_days, "night"))
        date_times = get_start_to_end_date_time(date_time)
        df_headers = []
        for date, time in date_times:
            header_df = f"{date}_{time}"
            val = self.get_value_from_df(header_df)
            if val in config.abbr_meal_off:
                df_headers.append(header_df)
        print(
            f"{config.warning_string} you are about to reset {len(df_headers)} meal cells for {self.name}"
        )
        confirmed = final_cofirm()
        if not confirmed:
            return
        self.reset_cell(df_headers)

    def reset_cell_guest_meal_date(self):
        date_time = self.select_date_and_time(show_text=f"reset date for {self.name}")
        if date_time is None:
            return
        date_times = get_start_to_end_date_time(date_time)
        df_headers = []
        for date, time in date_times:
            for header in self.boarder_values_headers:
                if "guest" in header and str(date) in header and str(time) in header:
                    df_headers.append(header)
        df_headers = list(set(df_headers))
        final_headers = []
        for header_df in df_headers:
            val = self.get_value_from_df(header_df)
            if int(val) == 0:
                final_headers.append(header_df)
        print(
            f"{config.warning_string} you are about to reset {len(final_headers)} guest-meal cells for {self.name}"
        )
        confirmed = final_cofirm()
        if not confirmed:
            return
        self.reset_cell(df_headers)

    def reset_cell_guest_meal_all(self):
        df_headers = []
        for header in self.boarder_values_headers:
            if "guest" in header:
                df_headers.append(header)
        final_headers = []
        for header_df in df_headers:
            val = self.get_value_from_df(header_df)
            if int(val) == 0:
                final_headers.append(header_df)
        print(
            f"{config.warning_string} you are about to reset {len(final_headers)} guest-meal cells for {self.name}"
        )
        confirmed = final_cofirm()
        if not confirmed:
            return
        self.reset_cell(df_headers)

    def reset_cell(self, df_header_list: list):
        for df_header in df_header_list:
            if df_header not in self.boarder_values_headers:
                print(
                    f"{config.failure_string}{df_header} is not present in barder dataframe."
                )
                continue
            cell = get_cell_for_boarder_from_df(
                self.boarder_values, self.name, df_header
            )
            write_in_excel_file(
                config.meal_count_sheet_name,
                cell,
                "",
                fill_reset=True,
                color_reset=True,
            )
            df_header = df_header.replace("_", "-")
            print(f"{config.success_string}done resetting cell: {cell} | {df_header}")

    def change_status(self):
        header = "Status"
        cell = get_cell_for_boarder_from_df(self.boarder_values, self.name, header)
        meal_status_txt = "off" if self.meal_status else "on"
        guest_meal_status_txt = "off" if self.guest_meal_status else "on"
        if cell in self.invalid_cells:
            return
        options = [
            f"turn {meal_status_txt} continue-meal",
            f"turn {guest_meal_status_txt} continue-guest-meal.",
        ]
        current_meal = "continue on" if self.meal_status else "continue off"
        current_guest_meal = "on" if self.guest_meal_status else "off"
        guest_meal_count_text = (
            f", count:{self.guest_meal_status_count}"
            if self.guest_meal_status_count > 0
            else ""
        )
        choice = select_items_from_list(
            options,
            multi_select=False,
            presentation_text=f"Status of {self.name}\nMeal: {current_meal} | Guest meal: {current_guest_meal}{guest_meal_count_text}",
        )
        if choice is None:
            return
        if choice == options[0]:
            return self.change_meal_status()
        if choice == options[1]:
            return self.change_guest_status()

    def change_guest_status(self):
        header = "Status"
        cell = get_cell_for_boarder_from_df(self.boarder_values, self.name, header)
        if cell in self.invalid_cells:
            return
        status_val = self.get_value_from_df(header)
        values = []
        if status_val:
            values = [val.strip() for val in status_val.split(",")]
        final_values = []
        current_guest_meal_count = self.guest_meal_status_count
        current_guest_meal_status = self.guest_meal_status
        nums = []
        for val in values:
            if val.isdigit():
                nums.append(val)
            if val not in config.abbr_continue_guest_meal:
                final_values.append(val)

        current_guest = (
            "guest continue on" if current_guest_meal_status else "guest continue off"
        )
        current_count = (
            f"(current count- {current_guest_meal_count})"
            if current_guest_meal_count > 1
            else ""
        )
        print(f"currently: {current_guest}, for {self.name}")
        meal_dialogue = "Change to: "
        guest_on_val = config.abbr_continue_guest_meal[0]
        if current_guest_meal_status:
            meal_dialogue += "continue guest meal off"
        else:
            meal_dialogue += "continue guest meal on"

        res = final_cofirm(meal_dialogue)
        if not res:
            return
        if res:
            if current_guest_meal_status:
                self.guest_meal_status = False
                self.guest_meal_status_count = 0
                for num in nums:
                    final_values.remove(num)
            if not current_guest_meal_status:
                final_values.append(guest_on_val)
                count = None
                while True:
                    count = input_digit_val(
                        f"Guest Meal count to apply {current_count}: "
                    )
                    if count is None:
                        return
                    if count < 1:
                        continue
                    final_values.append(str(count))
                    break
                confirmed = final_cofirm()
                if not confirmed:
                    return
                self.guest_meal_status_count = count
                self.guest_meal_status = True
                print("guest meal status:", self.guest_meal_status)
            # add meal on in final values
            meal_on_present = False
            meal_off_present = False
            for val in config.abbr_meal_on:
                if val in final_values:
                    meal_on_present = True
                    break
            for val in config.abbr_meal_off:
                if val in final_values:
                    meal_on_present = True
                    break
            if not meal_on_present and not meal_off_present:
                final_values.append(config.abbr_meal_on[0])

            final_values = ", ".join(final_values)
            write_in_excel_file(config.meal_count_sheet_name, cell, final_values)
            self.set_value_to_df(header, final_values)
            print(f"{config.success_string}successfully updated.")
            return True

    def change_meal_status(self):
        header = "Status"
        cell = get_cell_for_boarder_from_df(self.boarder_values, self.name, header)
        if cell in self.invalid_cells:
            return
        status_val = self.get_value_from_df(header)
        current_meal = "continue on" if self.meal_status else "continue off"
        print(f"currently: {current_meal}, for {self.name}")
        meal_dialogue = "Change to: "
        meal_on_val = config.abbr_meal_on[0]
        meal_off_val = config.abbr_meal_off[0]
        if self.meal_status:
            meal_dialogue += "meal continue off"
        else:
            meal_dialogue += "meal continue on"

        res = final_cofirm(meal_dialogue)
        if not res:
            return
        if res:
            final_values = []
            values = []
            if status_val:
                values = [val.strip() for val in status_val.split(",")]
            for val in values:
                if val not in config.abbr_meal_on + config.abbr_meal_off:
                    final_values.append(val)
            if self.meal_status:
                final_values.append(meal_off_val)
                self.meal_status = False
            elif not self.meal_status:
                final_values.append(meal_on_val)
                self.meal_status = True
            final_values = " ,".join(final_values)
            write_in_excel_file(config.meal_count_sheet_name, cell, final_values)
            self.set_value_to_df(header, final_values)
            print(f"{config.success_string}successfully updated.")
            return True

    ############## Read Data
    def get_deposit(self):
        header = "Deposit"
        cell = get_cell_for_boarder_from_df(self.boarder_values, self.name, header)
        if cell in self.invalid_cells:
            return None
        val = self.get_value_from_df(header)
        if not val:
            return 0
        return val

    def get_accessories_charge(self):
        header = "Acs Charge"
        cell = get_cell_for_boarder_from_df(self.boarder_values, self.name, header)
        if cell in self.invalid_cells:
            return None
        val = self.get_value_from_df(header)
        if not val:
            return 0
        return val

    def get_total_meal(self):
        headers = []
        for header in self.boarder_values_headers:
            if header.endswith("_night") or header.endswith("_day"):
                headers.append(header)
                cell = get_cell_for_boarder_from_df(
                    self.boarder_values, self.name, header
                )
                if cell in self.invalid_cells:
                    return None

        count = 0
        for header in headers:
            val = self.get_value_from_df(header)
            if val in config.abbr_meal_on:
                count += 1
        return count

    def get_total_meal_today(self):
        respect_year = config.respect_year_for_meal_counting
        respect_month = config.respect_month_for_meal_counting
        calculate = True
        if respect_year:
            if self.year != pendulum.now().year:
                calculate = False
        if respect_month:
            if pendulum.now().month != self.month:
                calculate = False

        if not calculate:
            return None
        current_date, current_time = get_date_time_now()
        dates = [i for i in range(1, current_date + 1)]
        headers = []
        for header in self.boarder_values_headers:
            if header.endswith("_night") or header.endswith("_day"):
                header_date = int(header.split("_")[0])
                if header_date in dates:
                    if header_date == current_date:
                        if current_time == "day":
                            if header.endswith("_night"):
                                continue
                    headers.append(header)
                    cell = get_cell_for_boarder_from_df(
                        self.boarder_values, self.name, header
                    )
                    if cell in self.invalid_cells:
                        return None

        count = 0
        for header in headers:
            val = self.get_value_from_df(header)
            if val in config.abbr_meal_on:
                count += 1
        return count

    def get_guest_meals(self):
        headers = []
        for header in self.boarder_values_headers:
            if "guest" in header:
                if "grand" not in header:
                    headers.append(header)
                    cell = get_cell_for_boarder_from_df(
                        self.boarder_values, self.name, header
                    )
                    if cell in self.invalid_cells:
                        return None
        guest_meals_df = self.boarder_values[headers]
        return guest_meals_df.sum().sum()

    def get_grand_guest_meals(self):
        headers = []
        for header in self.boarder_values_headers:
            if "guest" in header:
                if "grand" in header:
                    headers.append(header)
                    cell = get_cell_for_boarder_from_df(
                        self.boarder_values, self.name, header
                    )
                    if cell in self.invalid_cells:
                        return None
        grand_guest_meals_df = self.boarder_values[headers]
        return grand_guest_meals_df.sum().sum()

    def get_total_eggs(self):
        headers = []
        for header in self.boarder_values_headers:
            if "egg" in header:
                headers.append(header)
                cell = get_cell_for_boarder_from_df(
                    self.boarder_values, self.name, header
                )
                if cell in self.invalid_cells:
                    return None
        eggs_df = self.boarder_values[headers]
        return eggs_df.sum().sum()

    ############## Helpers
    def make_df_label(self, date: int, original_label: str) -> str:
        return f"{date}_{original_label}"

    def get_value_from_df(self, header, nan_val=""):
        index = self.boarder_values_headers.index(header)
        val = self.boarder_values.iloc[0, index]
        if isinstance(val, str):
            return val
        if math.isnan(val):
            return nan_val
        return val

    def set_value_to_df(self, header, val):
        index = self.boarder_values_headers.index(header)
        self.boarder_values.iloc[0, index] = val

    def select_date_and_time(
        self,
        month=None,
        year=None,
        show_tomorrow=True,
        show_yesterday=True,
        show_date_to_date=True,
        show_text="date",
    ):
        if month is None:
            month = self.month
        if year is None:
            year = self.year
        date = None
        date_today = int(pendulum.today().format("DD"))
        day_today = pendulum.today().format("dddd")
        select_date_str = "choose a date"
        select_date_to_date_str = "Date to Date"
        today_text = f"today ({day_today})"
        tomorrow_text = "tomorrow"
        yesterday_text = "yesterday"
        curr_month = pendulum.now().month
        curr_year = pendulum.now().year
        options = []
        if month == curr_month and year == curr_year:
            options.append(today_text)
            if show_tomorrow and date_today < self.month_days - 1:
                options.append(tomorrow_text)
            if show_yesterday and not date_today < 2:
                options.append(yesterday_text)
        options.append(select_date_str)
        if show_date_to_date:
            options.append(select_date_to_date_str)
        choice = select_items_from_list(
            options,
            multi_select=False,
            quit_text="CANCEL",
            presentation_text=f"Choose a {show_text} | today: {date_today}, {self.month_string.title()}",
        )
        if choice is None:
            return
        if choice == today_text:
            date = date_today
        if choice == tomorrow_text:
            date = int(pendulum.tomorrow().format("DD"))
        if choice == yesterday_text:
            date = int(pendulum.yesterday().format("DD"))
        if choice == select_date_str:
            date = choose_a_date(
                self.month_days,
                self.month_string,
                default=int(pendulum.today().format("DD")),
            )
        if choice == select_date_to_date_str:
            return self.select_date_to_date()
        if date is None:
            return

        time = None
        time = choose_time()
        if time is None:
            return

        return (date, time)

    def select_date_to_date(self):
        date_today = int(pendulum.today().format("DD"))
        day_today = pendulum.today().format("dddd")
        from_date = choose_a_date(
            self.month_days,
            self.month_string,
            default=date_today,
            show_text=f"START Date | today: {date_today}, {day_today}",
        )
        if from_date is None:
            return
        choose_time_str = "start time:"
        from_time = choose_time(show_text=choose_time_str)
        if from_time is None:
            return

        to_date = choose_a_date(
            self.month_days,
            self.month_string,
            default=self.month_days,
            show_text=f"END Date | today: {date_today}, {day_today}",
        )
        if to_date is None:
            return
        choose_time_str = "end time:"
        to_time = choose_time(show_text=choose_time_str)
        if to_time is None:
            return
        return ((from_date, from_time), (to_date, to_time))
