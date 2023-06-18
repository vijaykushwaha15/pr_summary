# pr_summary
GitHub API to retrieve a summary of all opened, closed, and in draft pull requests

# Note : Make sure to replace the 'fromaddr' and 'toaddr' variables with your email addresses and 'msecret' with your email password.

Feel free to modify the script according to your specific requirements.

This Python script retrieves pull requests from a GitHub repository, generates a summary, and sends it via email with an attached Excel file containing the details.
Prerequisites

Before running the script, ensure that you have the following dependencies installed:

    smtplib
    requests
    datetime
    xlsxwriter

You can install these dependencies using pip:

pip3 install smtplib requests xlsxwriter

# Usage

    Clone the repository:

bash

git clone https://github.com/vijaykushwaha15/pr_summary.git

    Navigate to the project directory:

bash

cd pr_summary

    Make sure to replace the following variables in the code with your desired values:

    base_url: The base URL of the GitHub API.
    repository: The repository in the format username/repository.

    Run the Python script:

python3 pull_requests_summary.py

Functionality:

The script performs the following tasks:

    Retrieves the pull requests from the specified repository created within the last week.
    Separates the pull requests into three categories: open, closed, and draft.
    Prints a summary of the pull requests in the console.
    Generates an Excel file named pull_requests.xlsx with three sheets: Open, Closed, and Draft.
    Populates the sheets with the pull request details, including the pull request number, username, PR summary, and timestamp.
    Sends an email with the Excel file attached to the specified recipient.

# Note : Make sure to replace the 'fromaddr' and 'toaddr' variables with your email addresses and 'msecret' with your email password.
