# Copilot Chat Adoption Insights


## Overview

This project contains a script designed to notify Microsoft BeLux V-teams about Copilot Chat adoptions for their accounts. The script automates specific notification tasks to streamline workflows and improve efficiency.

## Purpose

- Gaining insights about Copilot Chat adoption and providing them to V-teams at Microsoft BeLux.
- Providing visual insights and analytics to support decision-making on how to place Copilot Chat within managed accounts.
- Simplifies the process of reaching many V-teams.

## Orchestration

1. Looks for a MS Edge Work Profile that has access to the Power BI AE dashboard and MW SharePoint folder
2. Gathers snapshots from the Power BI accounts dashboard as csv and tries to update the stored csv on the SharePoint
3. Loops through the SharePoint csv of a V-team identified by ID
4. Screenshots visuals from the Power BI of a specific ID
5. Filters the stored snapshots data for a specific ID
6. Uses stored snapshots to build a graph to visualise the adoption progress for the accounts of a V-team
7. Accumulates visuals for those that prefer to be mailed once and not included on any CC
8. Reads the SharePoint csv to structure the mail and send it out to a V-team and people that want to be mailed only once
9. Reads SharePoint to skip those that do not want to be mailed whatsoever
10. Reads `.env` to attach additional information to a mail

## Run fails
- If the run fails, run again
- The run will take into account previous logs and avoid mailing again the same people

## Env variables
Create an `.env` file in the project's root, then write the following variables:
- `default_emails`: additional emails to append to the CCs
- `test_cc`: the CC used during test runs
- `test_to`: the To used during test runs
- `test_excluded`: list of people that will be notified once and excluded from CCs
- `test_excluded_names`: name of people that will be notified once and excluded from CCs. **Note:** Divided by `, `
- `url_sharepoint`: url to the SharePoint with relevant `txt` data in *csv format*
- `url_powerbi_export`: url to the Power BI dashboard with tables with columns to export
- `url_powerbi_show`: url to the Power BI dashboard with visuals to show

## SharePoint folder files

- `Excluded from CCs.txt`
  - Columns: FullName, Email, V-team role
  - List of people that will be mailed once and excluded from CCs
- `Filter accounts.txt`
  - Columns: ID, Accounts
  - Filters accounts for a specific V-team to consider when sharing the accounts overview visual
  - Separate *Accounts* by `|` 
- `Mail structure - AE.txt` or `Mail structure - Excluded.txt`
  - Columns: Command, Text
  - Outlines how a mail should be composed
  - AE: outline for Account Execs
  - Excluded: outline for people that will be mailed once
  - Column *Text* is the text content that will be written on the mail body
  - Column *Command*:
    - **Note**: Commands can be run multiple times simultaneously (E.g. of a command "break line", or "bullet line", or "break title", etc...)
    - *subject* => Subject of a mail
    - *line* => Writes paragragh
    - *title* => Writes bold paragragh
    - *break* => Inserts a break
    - *bullet* => Inserts bullets
    - *visual* => inserts a visual. The id of the visual is anything after a colon (E.g. visual: This is an ID;Weekly progress of the Copilot Chat MAU among your Accounts:)
    - *save* => If coupled with command *visual*, it will be stored to be later sent to an exclusion
    - *visuals* => It releases all the accumulated visuals for an exclusion
    - *url* => inserts a url. The placeholder of the url is anything after a colon (E.g. url: This is a placeholder;www.webaddress.com)
    - *end* => inserts a paragrah that will end the composing 
- `Skip.txt`
  - Columns: Email, ID
  - Avoids mailing a member in a specified V-team  
- `Table Data.txt`
  - Columns: TopParent, Copilot Chat MAU (Paid +UnPaid), Copilot Chat MAU (Paid +UnPaid)/TAM, Copilot Chat Unpaid MAU, Copilot Chat Paid MAU, Copilot Chat H2 Incremental MAU, Total TAM, AE, Date
  - Copilot Chat adoption data
- `V-teams.txt`
  - Columns: ID, Role, Names, To, CC
  - List of V-teams to mail
  - Separate *Names* values by `,`
  - Separate *To* and *CC* values by `whitespace`
- `Actual TAM.txt`
  - Columns: ID, Accounts, TAM
  - List of Accounts and their actual TAM
  - Separate *Accounts* by `|` 

## Logging

- In case of issues, a `logs.txt` file will describe it
- In case of issues with mailing an AE, a `mail_logs.txt` file will show who has been mailed
- During test runs, `test_logs.txt` and `test_mail_logs.txt` will be created instead

## Timing
- The script should be run 7 or more days after the las time:
- The Power BI dashboard shows data taken the day before
- Recommended to let the script run in the background
  - It can take up to 3 min to gather visuals and compose a single mail
  - Should be run while working on something else 

## Visuals
- The html attributes are indicated in dictionaries the `browser.py` file in `/vars`

## Requirements

- Git
- Python 3.8 or higher
- Microsoft Edge installed
- *MSFT-AzVPN-Manual* turned on
- Run it in a terminal opened **as administrator**
- Required Python libraries (see `requirements.txt`)
- Working Profile folder in Microsoft Edge's appdata folder

## Installation

1. Clone the repository:
  ```bash
  git clone <repository-url>
  ```
2. Navigate to the project directory:
  ```bash
  cd Copilot-Chat-Adoption-Insights
  ```
3. Install dependencies:
  ```bash
  pip install -r requirements.txt
  ```

## Usage

Run the script using the following command:
```bash
python main.py
```

Follow the instructions on the terminal to interact with the tool.
