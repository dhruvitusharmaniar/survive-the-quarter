# Name: Dhruvi Maniar
# Group: C
# Project: Survive the Quarter (DESMI 2026, Python for Digital Marketing, M1)
#
# ---------------------------------------------------------------------------
# 1. WHAT GOES IN?
# My scenario file is scenarios_C.csv. The separator is a comma. Text that
# contains a comma or quotes is wrapped in double quotes, so I read it with
# the csv module instead of split(",").
# The first line is a header with 14 columns:
#   Event                         -> the situation of the day
#   Day                           -> the scenario number (1 to 33)
#   Decision 1, 2, 3              -> the text of options A, B and C
#   Reputation 1, 2, 3            -> reputation impact, written like "+15 pts"
#   Budget 1, 2, 3                -> budget impact, written like "-800 EUR"
#   Outcome 1, 2, 3               -> the feedback line shown after the choice
# The numbers are text with a sign and a unit, so I remove " pts" or " EUR"
# and turn what is left into an int. In my file the columns are grouped by
# type, not by option: the impact of option B is in "Decision 2",
# "Budget 2", "Reputation 2" and "Outcome 2".
#
# 2. WHAT HAPPENS? (one day, in order)
#   - show the day number and both gauges (budget and reputation)
#   - show the situation and options A, B, C
#   - ask for A, B or C until the answer is valid (lowercase accepted)
#   - apply the budget and reputation change, reputation capped at 100
#   - show the feedback line and the changes
#   - save the day in the history list
#   - if budget or reputation is 0 or less, the game stops at once
#
# 3. WHAT COMES OUT?
# The player sees: fired (and why) or "Quarter survived", the budget left,
# the reputation, the final score and the title, then the best and the worst
# move of the game, and whether the high score (best budget) was beaten.
# One line is appended to results.csv, for example:
#   Dhruvi,6350,100,CMO Material,2026-10-30 14:05
#
# 4. WHAT COULD GO WRONG?
#   - The scenario file is missing: I print a clear message and stop
#     without a crash.
#   - The game is started from another folder (for example with the Run
#     button of an editor): I look for the files next to the script, not in
#     the current folder, so it still works.
#   - The scenario file was opened and saved again in Excel: Excel can add
#     an invisible mark at the start (BOM), use ";" instead of ",", or change
#     the encoding. I handle all three, and if the file cannot be read
#     at all I print a message instead of crashing.
#   - A line of the file is broken (wrong number of columns or a number
#     I cannot read): I skip that line and say so, the game goes on.
#   - The player types something other than A, B or C: I ask again.
#   - The player types nothing for the name: I use "Player".
#   - The name contains invisible characters (odd key combinations): I
#     remove them, because some Python versions crash when they write or
#     read them in a CSV file.
#   - The player presses Ctrl+D or Ctrl+C: I say goodbye and stop cleanly.
#   - A gauge goes below zero: the player is fired right away.
#   - Reputation goes above 100: I put it back to 100.
#   - The file has fewer than 10 scenarios: I play as many as there are.
#   - results.csv does not exist yet: I create it with a header line, and
#     the high score feature says it is the first game.
#   - results.csv is open in Excel (locked) or was damaged: the game still
#     ends normally and tells the player the result could not be saved.
#   - A fired player has a big budget left: only games that survived
#     count for the high score.
#
# FEATURE CHOSEN: c) High score
# At the end, the game reads results.csv and says whether the player beat
# the best budget of all the previous games that survived the quarter.
# Why: it makes me read a CSV file a second time, but a file my own program
# wrote, and it gives players a reason to play again and beat each other.
# ---------------------------------------------------------------------------

import csv
import io
import os
import random
import sys
from datetime import datetime

# Game settings, taken from the brief
SCENARIO_NAME = "scenarios_C.csv"
RESULTS_NAME = "results.csv"
START_BUDGET = 10000
START_REPUTATION = 50
MAX_REPUTATION = 100
DAYS_PER_GAME = 10
LETTERS = ("A", "B", "C")

# Look for the files in the folder of this script, so the game works even
# when it is started from another folder
SCRIPT_FOLDER = os.path.dirname(os.path.abspath(__file__))
SCENARIO_FILE = os.path.join(SCRIPT_FOLDER, SCENARIO_NAME)
RESULTS_FILE = os.path.join(SCRIPT_FOLDER, RESULTS_NAME)


# Make sure a strange character typed or printed never crashes the game:
# a character the terminal cannot show is replaced by "?"
def protect_terminal():
    try:
        sys.stdout.reconfigure(errors="replace")
        sys.stdin.reconfigure(errors="replace")
    except (AttributeError, ValueError):
        pass  # very old Python or unusual terminal: keep the default


# Read a line from the keyboard. If the player presses Ctrl+D or Ctrl+C,
# stop the game cleanly instead of crashing.
def read_line(prompt):
    try:
        return input(prompt)
    except (EOFError, KeyboardInterrupt):
        print("\nGame interrupted. See you next quarter!")
        sys.exit(0)


# Ask the player's name. Invisible characters (from odd key combinations)
# are removed, the name is cut to 30 letters, and an empty answer
# becomes "Player".
def ask_name():
    typed = read_line("What is your name? ")
    name = ""
    for character in typed:
        if character.isprintable():
            name = name + character
    name = name.strip()[:30]
    if name == "":
        name = "Player"
    return name


# Turn a text like "+15 pts" or "-800 EUR" into a number (15 or -800).
def parse_number(text):
    clean = text.replace("pts", "").replace("EUR", "").strip()
    return int(clean)


# Read the scenario file and return a list of scenarios.
# Each scenario is a dictionary, and each option inside it is a tuple:
# (text, budget_change, reputation_change, feedback)
def load_scenarios(filename):
    scenarios = []
    skipped = 0
    # "utf-8-sig" also accepts the invisible mark Excel adds at the start,
    # and errors="replace" turns a character saved in another encoding
    # into "?" instead of crashing
    with open(filename, newline="", encoding="utf-8-sig", errors="replace") as file:
        # Remove invisible "null" characters: old Pythons refuse to read them
        content = file.read().replace("\x00", "")
    # My file uses ",", but Excel in French saves with ";"
    lines = content.splitlines()
    first_line = ""
    if len(lines) > 0:
        first_line = lines[0]
    separator = ","
    if first_line.count(";") > first_line.count(","):
        separator = ";"
    reader = csv.DictReader(io.StringIO(content, newline=""), delimiter=separator)
    for row in reader:
        try:
            options = {}
            for number, letter in enumerate(LETTERS, start=1):
                text = row["Decision " + str(number)].strip()
                budget_change = parse_number(row["Budget " + str(number)])
                reputation_change = parse_number(row["Reputation " + str(number)])
                feedback = row["Outcome " + str(number)].strip()
                options[letter] = (text, budget_change, reputation_change, feedback)
            scenarios.append({"event": row["Event"].strip(), "options": options})
        except (KeyError, ValueError, TypeError, AttributeError):
            skipped = skipped + 1
    if skipped > 0:
        print("Warning:", skipped, "broken line(s) skipped in", SCENARIO_NAME)
    return scenarios


# Show the day number, both gauges, the situation and the three options.
def show_day(day_number, budget, reputation, scenario):
    print()
    print("--- Day", day_number, "---")
    print("[Budget:", budget, "EUR | Reputation:", str(reputation) + "/" + str(MAX_REPUTATION) + "]")
    print(scenario["event"])
    for letter in LETTERS:
        print(" " + letter + ":", scenario["options"][letter][0])


# Ask for A, B or C until the answer is valid. Lowercase is accepted.
def ask_choice():
    answer = read_line("Your choice (A, B or C): ").strip().upper()
    while answer not in LETTERS:
        answer = read_line("Please type A, B or C: ").strip().upper()
    return answer


# Apply one option to the gauges and return the new gauges as a tuple.
# Reputation can never go above the maximum.
def apply_choice(budget, reputation, option):
    text, budget_change, reputation_change, feedback = option
    budget = budget + budget_change
    reputation = min(reputation + reputation_change, MAX_REPUTATION)
    return budget, reputation


# Format a change with its sign, for example +20 or -2000.
def signed(number):
    if number >= 0:
        return "+" + str(number)
    return str(number)


# Return the reason the player is fired, or an empty text if still alive.
def fired_reason(budget, reputation):
    if budget <= 0 and reputation <= 0:
        return "the budget and the reputation both hit zero"
    if budget <= 0:
        return "the budget hit zero, the bank says no"
    if reputation <= 0:
        return "the reputation hit zero, the internet hates you"
    return ""


# Compute the final score and the title. Returns a tuple (score, title).
def get_score_and_title(budget, reputation, fired):
    if fired:
        return 0, "Fired"
    score = reputation + budget // 200
    if score >= 140:
        title = "Viral Legend"
    elif score >= 120:
        title = "CMO Material"
    elif score >= 95:
        title = "Growth Hacker"
    else:
        title = "Eternal Intern"
    return score, title


# Find the best and the worst move in the history.
# Each history entry is a tuple (reputation_change, day, text), so max()
# and min() compare the reputation change first.
def find_best_and_worst(history):
    return max(history), min(history)


# Feature c: read results.csv and return the best budget of the previous
# games that survived, or None if there is no such game yet.
# Fired games do not count: you cannot hold a record after being fired.
# If the file is missing, locked or damaged, there is simply no record.
def read_best_budget(filename):
    best = None
    try:
        with open(filename, newline="", encoding="utf-8-sig", errors="replace") as file:
            content = file.read().replace("\x00", "")
        for row in csv.reader(io.StringIO(content, newline="")):
            if len(row) < 4 or row[3].strip() == "Fired":
                continue
            try:
                budget = int(row[1])
            except ValueError:
                continue  # the header line, or a broken line
            if best is None or budget > best:
                best = budget
    except (OSError, csv.Error):
        return None
    return best


# Add one line to results.csv: name, budget, reputation, title, date.
# Write the header first if the file is new or empty, and start on a new
# line if someone edited the file and removed the last line break.
# Returns True if the line was saved, False if the file could not be written.
def save_result(filename, name, budget, reputation, title):
    try:
        with open(filename, "r", encoding="utf-8", errors="replace") as file:
            old_content = file.read()
    except FileNotFoundError:
        old_content = ""
    except OSError:
        return False
    date = datetime.now().strftime("%Y-%m-%d %H:%M")
    try:
        with open(filename, "a", newline="", encoding="utf-8") as file:
            if old_content.strip() != "" and not old_content.endswith("\n"):
                file.write("\r\n")
            writer = csv.writer(file)
            if old_content.strip() == "":
                writer.writerow(["Name", "Budget", "Reputation", "Title", "Date"])
            writer.writerow([name, budget, reputation, title, date])
    except (OSError, csv.Error):
        return False
    return True


# The whole game, from the name to the saved result.
def main():
    protect_terminal()
    print("=" * 30)
    print("SURVIVE THE QUARTER")
    print("Croissant Rocket: croissants delivered by drone.")
    print("Keep budget and reputation above zero for", DAYS_PER_GAME, "days.")
    print("=" * 30)

    # Load the scenarios, stop cleanly if the file is missing or unreadable
    try:
        scenarios = load_scenarios(SCENARIO_FILE)
    except FileNotFoundError:
        print("Error:", SCENARIO_NAME, "was not found.")
        print("Put it in the same folder as this script:", SCRIPT_FOLDER)
        return
    except (OSError, UnicodeDecodeError, csv.Error):
        print("Error:", SCENARIO_NAME, "could not be read.")
        print("Close it if it is open in Excel, or put back the original file.")
        return
    if len(scenarios) == 0:
        print("Error: no valid scenario in", SCENARIO_NAME)
        print("Put back the original file (do not save it from Excel).")
        return

    name = ask_name()

    # Shuffle and keep 10 days (or fewer if the file is short)
    random.shuffle(scenarios)
    days = scenarios[:DAYS_PER_GAME]

    budget = START_BUDGET
    reputation = START_REPUTATION
    history = []
    reason = ""

    # Play the days one by one
    for day_number, scenario in enumerate(days, start=1):
        show_day(day_number, budget, reputation, scenario)
        letter = ask_choice()
        option = scenario["options"][letter]
        budget, reputation = apply_choice(budget, reputation, option)

        text, budget_change, reputation_change, feedback = option
        print(feedback)
        print("Budget", signed(budget_change), "EUR, Reputation", signed(reputation_change))
        history.append((reputation_change, day_number, text))

        # A gauge at 0 or less ends the game at once
        reason = fired_reason(budget, reputation)
        if reason != "":
            break

    # Show the result
    fired = reason != ""
    score, title = get_score_and_title(budget, reputation, fired)
    print()
    print("=" * 30)
    if fired:
        print("You are fired!", name + ",", reason + ".")
        print("Budget left:", budget, "EUR, reputation:", str(reputation) + "/" + str(MAX_REPUTATION))
    else:
        print("Quarter survived!")
        print("Budget left:", budget, "EUR, reputation:", str(reputation) + "/" + str(MAX_REPUTATION))
        print("Final score:", score)
    print(name + ", your title:", title)

    # Best and worst move, using the history of tuples
    best, worst = find_best_and_worst(history)
    print("Best move: day", best[1], '"' + best[2] + '"', "(" + signed(best[0]) + ")")
    print("Worst move: day", worst[1], '"' + worst[2] + '"', "(" + signed(worst[0]) + ")")

    # Feature c: compare with the best budget saved so far
    best_budget = read_best_budget(RESULTS_FILE)
    if fired:
        print("High score: fired managers cannot take the record.")
    elif best_budget is None:
        print("High score: first game on record, you set the bar at", budget, "EUR.")
    elif budget > best_budget:
        print("New high score! You beat the best budget of", best_budget, "EUR.")
    else:
        print("High score not beaten. The best budget is still", best_budget, "EUR.")

    # Save one line in results.csv
    if save_result(RESULTS_FILE, name, budget, reputation, title):
        print("Saved in", RESULTS_NAME)
    else:
        print("Could not save in", RESULTS_NAME + ". Close it if it is open in Excel, then play again.")


if __name__ == "__main__":
    main()
