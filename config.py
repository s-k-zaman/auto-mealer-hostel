workbook_name = "hostel-meals.xlsx"
meal_count_sheet_name = "meal_counts"
marketing_sheet_name = "marketings"
establishment_sheet_name = "establishment"
meal_routine_sheet_name = "meal_combination"


## all these used in chaecking ivalid data.
## should use value from the lists
# rules
abbr_no_beef = ["nb", "no beef"]
abbr_no_fish = ["nf", "no fish"]
abbr_no_egg = ["ne", "no egg"]
abbr_no_chicken = ["nc", "no chicken"]
abbr_only_night = ["on", "only night"]
abbr_only_day = ["od", "only day"]
abbr_weekend_off = ["wo", "weekend off"]  # saturday, sunday
abbr_weekend_on = ["won", "weekend on"]  # saturday, sunday
abbr_saturday_night_off = ["sanoff", "saturday night off"]
abbr_sunday_off = ["soff", "sunday off"]
abbr_sunday_on = ["son", "sunday on"]
# status
abbr_status_on = ["running", "on", "yes"]
abbr_status_off = ["stopped", "off", "no"]
abbr_continue_guest_meal = [
    "guest",
    "cgm",
]  # [.. , amount] amount -> no. of guest meal to on
# meal on
abbr_meal_on = ["on", "yes", "true"]
abbr_meal_off = ["off", "no", "false"]

success_string = " ✅ "
failure_string = " ❌ "
info_string = " ℹ️  "
warning_string = " ❗ "
selected_string = " 🔵 "
meal_on_string = " 🟢 "
meal_off_string = " 🔴 "


allow_all_boarder_selection = True

## times use 24 hr format
## IMPORTANT: always use string for time
count_night_meal_after = "14.30"
count_day_meal_after = "21"
respect_month_for_meal_counting = True
respect_year_for_meal_counting = True

rupee_symbol = "₹"
invalid_text = "Invalid value present"
MANAGER_NAME = "Khairul"

routine = [
    ["weekday", "day", "night"],
    ["sunday", "chicken", "veg"],
    ["monday", "fish", "egg"],
    ["tuesday", "egg", "beef"],
    ["wednesday", "veg", "fish"],
    ["thursday", "egg", "chicken"],
    ["friday", "fish", "egg"],
    ["saturday", "veg", "beef"],
]

routine_exp = [
    ["weekday", "day", "night"],
    ["sunday", "chicken", "veg"],
    ["monday", "fish", "egg"],
    ["tuesday", "egg", "beef"],
    ["wednesday", "veg", "fish"],
    ["thursday", "egg", "chicken"],
    ["friday", "fish", "egg"],
    ["saturday", "veg", "beef"],
]

guest_meal_price = 45
grand_guest_price = 128
egg_price = 10

approx_meal_charge = 35
approx_establishment_charge = 400
threshold_amount_to_meal_off = 100
thresold_meal_count_for_warning = 2


# enable if updatig codes.
debugging = False
