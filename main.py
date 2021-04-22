import json
import subprocess
import re
import time
from datetime import datetime
from threading import Timer
from typing import Optional, List

import selenium.webdriver.remote.webelement
from selenium import webdriver
from selenium.common import exceptions
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager


class Team:
    def __init__(self, name, team_id, channels=None):
        self.name = name
        self.team_id = team_id

        if channels is None:
            self.getChannels()
        else:
            self.channels = channels

    def __str__(self):
        channel_string = '\n\t'.join([str(channel) for channel in self.channels])
        return f"{self.name}\n\t{channel_string}"

    def getElement(self):
        team_header = browser.find_element_by_css_selector(f"h3[id='{self.team_id}'")
        team_element = team_header.find_element_by_xpath("..")
        return team_element

    def expandChannels(self):
        try:
            self.getElement().find_element_by_css_selector("div.channels")
        except exceptions.NoSuchElementException:
            try:
                self.getElement().click()
                self.getElement().find_element_by_css_selector("div.channels")
            except (exceptions.NoSuchElementException, exceptions.ElementNotInteractableException,
                    exceptions.ElementClickInterceptedException):
                return None

    def getChannels(self):
        self.expandChannels()
        channels = self.getElement().find_elements_by_css_selector(".channels > ul > ng-include > li")

        channel_names = [channel.get_attribute("data-tid") for channel in channels]
        channel_names = [channel_name[channel_name.find("channel-") + 8:channel_name.find("-li")] for channel_name in
                         channel_names]

        channel_ids = [channel.get_attribute("id").replace("channel-", "") for channel in channels]

        meeting_states = []
        for channel in channels:
            try:
                channel.find_element_by_css_selector("a > active-calls-counter")
                meeting_states.append(True)
            except exceptions.NoSuchElementException:
                meeting_states.append(False)

        self.channels = [Channel(channel_names[i], channel_ids[i], meeting_states[i]) for i in
                         range(len(channel_names))]


class Channel:
    def __init__(self, name, channel_id, has_meeting=False):
        self.name = name
        self.channel_id = channel_id
        self.has_meeting = has_meeting

    def __str__(self):
        return self.name + (" [MEETING]" if self.has_meeting else "")


class Meeting:
    def __init__(self, title, meeting_id, time_started, channel_id=None):
        self.title = title
        self.meeting_id = meeting_id
        self.time_started = time_started
        self.channel_id = channel_id

    def __str__(self):
        return f"\t{self.title} {self.time_started} [CHANNEL]"


# Global variable initialization
browser: Optional[webdriver.Chrome] = None
total_members: Optional[int] = None
config: Optional[dict] = None
meetings: List[Meeting] = []
current_meeting: Optional[Meeting] = None
already_joined_ids: List[str] = []
active_correlation_id: str = ""
hangup_thread: Optional[Timer] = None
conversation_link: str = "https://teams.microsoft.com/_#/conversations/a"
uuid_regex: str = r"\b[0-9a-f]{8}\b-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-\b[0-9a-f]\b"


def startRecording():
    """A method to start recording"""
    subprocess.run(["resources\\OBSCommand.exe", "/startrecording"])


def stopRecording():
    """A method to stop recording"""
    subprocess.run(["resources\\OBSCommand.exe", "/stoprecording"])


def loadConfig():
    """
    Loads the configuration file to set up the global variables
    """
    global config
    with open("config.json", encoding="utf-8") as json_data_file:
        config = json.load(json_data_file)


def initializeBrowser():
    """
    Initializes the webdriver browser
    """
    global browser

    browser_options = webdriver.ChromeOptions()
    browser_options.add_argument("--ignore-certificate-errors")
    browser_options.add_argument("--ignore-ssl-errors")
    browser_options.add_argument("--use-fake-ui-for-media-stream")
    browser_options.add_argument("--no-sandbox")
    browser_options.add_experimental_option(
        "prefs", {
            'credentials_enable_service': False,
            'profile.default_content_setting_values.media_stream_mic': 1,
            'profile.default_content_setting_values.media_stream_camera': 1,
            'profile.default_content_setting_values.geolocation': 1,
            'profile.default_content_setting_values.notifications': 1,
            'profile': {
                'password_manager_enabled': False
            }
        }
    )
    browser_options.add_experimental_option('excludeSwitches', ['enable-automation'])

    browser = webdriver.Chrome(ChromeDriverManager().install(), options=browser_options)
    browser.maximize_window()


def waitUntilFound(selector: str, timeout: int,
                   print_error: bool = True) -> Optional[selenium.webdriver.remote.webelement.WebElement]:
    """
    A method to make the webdriver wait until a specific element appears on the site by looking for its CSS selector. If
    the method finds the element it returns it otherwise it returns nothing.

    :param selector: The CSS selector of the element you wish to be present.
    :param timeout: Tells the method how long it should search for the element before giving up.
    :param print_error: Tells the method if it should print that there was a problem locating the element, provided
                        there was a problem.
    :return: The element attached to the selector
    """
    try:
        element_present = ec.visibility_of_element_located((By.CSS_SELECTOR, selector))
        WebDriverWait(browser, timeout).until(element_present)

        return browser.find_element_by_css_selector(selector)
    except exceptions.TimeoutException:
        if print_error:
            print(f"Timeout waiting for element: {selector}")
        return None


def switchToTeamsTab():
    """
    A method to find the teams button and then click it
    """
    teams_button = waitUntilFound("button.app-bar-link > ng-include > svg.icons-teams", 5)
    if teams_button is not None:
        teams_button.click()


def preparePage():
    try:
        browser.execute_script("document.getElementById('toast-container').remove()")
    except exceptions.JavascriptException:
        pass


def getAllTeams() -> List[Team]:
    """
    A method to find every team the user is a part of.

    :return: A list with every team in the page.
    """
    teams_elements = browser.find_elements_by_css_selector("ul>li[role='treeitem']>div[sv-element]")

    team_names = [teams_element.get_attribute("data-tid") for teams_element in teams_elements]
    team_names = [team_name[team_name.find("team-") + 5:team_name.find("-li")] for team_name in team_names]

    team_headers = [teams_element.find_element_by_css_selector("h3") for teams_element in teams_elements]
    team_ids = [team_header.get_attribute("id") for team_header in team_headers]

    return [Team(team_names[i], team_ids[i]) for i in range(len(teams_elements))]


def getMeetings(teams: List[Team]):
    """
    A method to update the list containing all the meetings based on the list of teams generated by 'getAllTeams()'.

    :param teams: The list of teams generated by 'getAllTeams()'.
    """
    global meetings

    for team in teams:
        for channel in team.channels:
            if channel.has_meeting:
                browser.execute_script(
                    f'window.location = "{conversation_link}a?threadID={channel.channel_id}&ctx=channel";')
                switchToTeamsTab()

                meeting_element = waitUntilFound(".ts-calling-thread-header", 10)
                if meeting_element is None:
                    continue

                meeting_elements = browser.find_elements_by_css_selector(".ts-calling-thread-header")
                for meeting_element in meeting_elements:
                    meeting_id = meeting_element.get_attribute("id")
                    time_started = int(meeting_id.replace("m", "")[:-3])

                    # Skip already joined meetings
                    correlation_id = meeting_element.find_element_by_css_selector(
                        "calling-join-button > button").get_attribute("id")
                    if active_correlation_id != "" and correlation_id.find(active_correlation_id) != -1:
                        continue

                    meetings.append(Meeting(f"{team.name} -> {channel.name}", meeting_id, time_started,
                                            channel_id=channel.channel_id))


def decideMeeting() -> Optional[Meeting]:
    """
    A method to return to decide which meeting should be joined.

    :return: The meeting that the method thinks should be joined.
    """
    global meetings
    newest_meetings = []

    if len(meetings) == 0:
        return None

    meetings.sort(key=lambda x: x.time_started, reverse=True)
    newest_time = meetings[0].time_started

    for meeting in meetings:
        if meeting.time_started >= newest_time:
            newest_meetings.append(meeting)
        else:
            break

    if (current_meeting is None or newest_meetings[0].time_started > current_meeting.time_started) and (
            current_meeting is None or newest_meetings[0].meeting_id != current_meeting.meeting_id) and (
            newest_meetings[0].meeting_id not in already_joined_ids):
        return newest_meetings[0]

    return None


def joinMeeting(meeting: Meeting):
    global hangup_thread, current_meeting, already_joined_ids, active_correlation_id

    hangup()
    browser.execute_script(f'window.location = "{conversation_link}a?threadId={meeting.channel_id}&ctx=channel";')
    switchToTeamsTab()

    join_button = waitUntilFound(f"div[id='{meeting.meeting_id}'] > calling-join-button > button", 5)
    if join_button is None:
        return None

    browser.execute_script("arguments[0].click()", join_button)

    join_now_button = waitUntilFound("button[data-tid='prejoin-join-button']", 30)
    if join_now_button is None:
        return None

    uuid = re.search(uuid_regex, join_now_button.get_attribute("track-data"))
    if uuid is not None:
        active_correlation_id = uuid.group(0)
    else:
        active_correlation_id = ""

    # Turn camera off
    video_button = browser.find_element_by_css_selector("toggle-button[data-tid='toggle-video'] > div > button")
    video_is_on = video_button.get_attribute("aria-pressed")
    if video_is_on == "true":
        video_button.click()
        print("The camera is now off")

    # Turn mic off
    mic_button = browser.find_element_by_css_selector("toggle-button[data-tid='toggle-mute'] > div > button")
    mic_is_on = mic_button.get_attribute("aria-pressed")
    if mic_is_on == "true":
        mic_button.click()
        print("The microphone is now off")

    # join_button = waitUntilFound(f"div[id='{meeting.meeting_id}'] > calling-join-button > button", 5)
    join_now_button = waitUntilFound("button[data-tid='prejoin-join-button']", 5)
    if join_now_button is None:
        return None

    join_now_button.click()

    current_meeting = meeting
    already_joined_ids.append(meeting.meeting_id)

    print(f"Joined meeting: {meeting.title}")
    startRecording()

    if "auto_leave_after_min" in config and config["auto_leave_after_min"] > 0:
        hangup_thread = Timer(config["auto_leave_after_min"] * 60, hangup)
        hangup_thread.start()


def getMeetingMembers() -> Optional[int]:
    """
    A method to count the members of the meeting.

    :return: The member count of the meeting.
    """
    meeting_elements = browser.find_elements_by_css_selector(".one-call")

    for meeting_element in meeting_elements:
        try:
            meeting_element.click()
            break
        except:
            continue

    time.sleep(2)

    try:
        browser.execute_script("document.getElementById('roster-button').click()")
        time.sleep(2)

        participants_element = browser.find_element_by_css_selector(
            "calling-roster-section[section-key='participantsInCall'] .roster-list-title")
        attendees_element = browser.find_element_by_css_selector(
            "calling-roster-section[section-key='attendeesInMeeting'] .roster-list-title")
    except (exceptions.JavascriptException, exceptions.NoSuchElementException):
        print("Failed to get meetings members")
        return None

    if participants_element is not None:
        participants = [int(s) for s in participants_element.get_attribute("aria-label").split() if s.isdigit()]
    else:
        participants = 0

    if attendees_element is not None:
        attendees = [int(s) for s in attendees_element.get_attribute("aria-label").split() if s.isdigit()]
    else:
        attendees = 0

    return sum(participants + attendees)


def hangup() -> Optional[bool]:
    """
    A method to determine if the meeting should be hung up on or not.

    :return: True if the meeting should be hang up on and False otherwise
    """
    global current_meeting, active_correlation_id

    if current_meeting is None:
        return None

    try:
        switchToTeamsTab()
        hangup_button = browser.find_element_by_css_selector("button[data-tid='call-hangup']")
        hangup_button.click()
        startRecording()

        current_meeting = None

        if hangup_thread:
            hangup_thread.cancel()

        return True
    except exceptions.NoSuchElementException:
        return False


def handleLeaveThreshold(current_members: int) -> bool:
    """
    A method to handle the logic of leaving a meeting based on the members.

    :param current_members: An int of how many members are in the meeting.
    :return: True if the meeting should be hung up on and False otherwise.
    """

    leave_number = config['leave_threshold_number']

    if leave_number == "":
        if 0 < current_members < 3:
            print("Last attendee in meeting")
            hangup()
            return True
    else:
        if current_members < float(leave_number):
            print("Last attendee in meeting")
            hangup()
            return True

    return False


def main():
    global config, meetings, conversation_link, total_members

    initializeBrowser()
    browser.get("https://teams.microsoft.com")

    if config['email'] != "" and config['password'] != "":
        login_email = waitUntilFound("input[type='email']", 30)
        if login_email is not None:
            login_email.send_keys(config['email'])
            login_email.send_keys(Keys.ENTER)

        login_user = waitUntilFound("input[type='text']", 30)
        if login_user is not None:
            login_user.send_keys(config['email'][:8])

        login_password = waitUntilFound("input[type='password']", 30)
        if login_password is not None:
            login_password.send_keys(config['password'])
            login_password.send_keys(Keys.ENTER)

        keep_logged_in = waitUntilFound("input[id='idBtn_Back']", 5)
        if keep_logged_in is not None:
            keep_logged_in.click()
        else:
            print("Login Unsuccessful, recheck entries in config")

        use_website = waitUntilFound(".use-app-link", 5, print_error=False)
        if use_website is not None:
            use_website.click()

    print("Waiting for correct page...", end='')
    if waitUntilFound("#teams-app-bar", 60*5) is None:
        exit(1)

    print("\rFound page, do not click anything on the webpage from now on.")
    time.sleep(5)

    preparePage()
    switchToTeamsTab()

    url = browser.current_url
    url = url[:url.find("conversations/") + 14]
    conversation_link = url

    check_interval = 10
    if "check_interval" in config and config["check_interval"] > 1:
        check_interval = config["check_interval"]

    interval_count = 0
    while True:
        timestamp = datetime.now()

        if "pause_search" in config and config["pause_search"] and current_meeting is not None:
            print(f"\n[{timestamp:%H:%M:%S}] Meeting search is paused because you are still in a meeting")
        else:
            print(f"\n[{timestamp:%H:%M:%S}] Looking for new meetings")

            switchToTeamsTab()
            teams = getAllTeams()

            if len(teams) == 0:
                print("Nothing found, is Teams in grid mode?")
                exit(1)
            else:
                getMeetings(teams)

            if len(meetings) > 0:
                print("Found meetings: ")
                for meeting in meetings:
                    print(meeting)

                meetings_to_join = decideMeeting()
                if meetings_to_join is not None:
                    total_members = 0
                    joinMeeting(meetings_to_join)

        meetings = []
        members_count = None
        if current_meeting is not None:
            members_count = getMeetingMembers()
            if members_count and members_count > total_members:
                total_members = members_count

        if "leave_if_last" in config and config["leave_if_last"] and interval_count % 5 == 0 and interval_count > 0:
            if current_meeting is not None and members_count is not None and total_members is not None:
                if handleLeaveThreshold(members_count):
                    total_members = None

        interval_count += 1
        time.sleep(check_interval)


if __name__ == "__main__":
    loadConfig()

    if "run_at_time" in config and config["run_at_time"] != "":
        now = datetime.now()
        run_at = datetime.strptime(config["run_at_time"], "%H:%M").replace(year=now.year, month=now.month, day=now.day)

        if run_at.time() < now.time():
            run_at = datetime.strptime(config["run_at_time"], "%H:%M").replace(year=now.year, month=now.month,
                                                                               day=now.day + 1)

        start_delay = (run_at - now).total_seconds()
        print(f"Waiting until {run_at} ({int(start_delay)}s)")
        time.sleep(start_delay)

    try:
        main()
    finally:
        if browser is not None:
            browser.quit()

        if hangup_thread is not None:
            hangup_thread.cancel()
