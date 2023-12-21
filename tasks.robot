*** Settings ***
Documentation     Template robot main suite.
# Library           RPA.Browser.Selenium
# Library           RPA.Email.ImapSmtp
# Library           CICProcess.py
Library           Bot.py
Task Teardown     Teardown

*** Tasks ***
Prime CIC Task
    # Setup Platform Components
    Start