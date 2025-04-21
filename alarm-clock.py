#!/usr/bin/env python3
import os
import sys
import json
import locale
import calendar
import subprocess
import signal
import requests
import urllib
import shutil
from time import sleep
from datetime import datetime, time, timezone, timedelta
from PyQt6.QtCore import Qt, QTimer, QTime, QSize
from PyQt6.QtGui import QIcon, QColor, QCursor
from PyQt6.QtWidgets import (
    QApplication,
    QMainWindow,
    QPushButton,
    QScrollArea,
    QHBoxLayout,
    QVBoxLayout,
    QSystemTrayIcon,
    QMenu,
    QFrame,
    QLabel,
    QWidget,
    QLineEdit,
    QTimeEdit,
    QCheckBox,
    QRadioButton,
    QGraphicsOpacityEffect,
    QMessageBox,
)


def createCustomFont(size=10, weight=400):
    font = app.font()
    font.setPointSize(size)
    font.setWeight(weight)
    return font


def getGrayColor():
    color = app.palette().text().color()
    color.setAlpha(128)
    return color.name(QColor.NameFormat.HexArgb)


class Alarm:
    def __init__(self):
        self.name = ""
        self.repeat = []  # Repeat is from 0-6 (Monday - Sunday)
        self.time = time(0, 0, 0)
        self.enabled = True


class OutlookReminder:
    def __init__(self):
        self.id = ""
        self.subject = ""
        self.location = ""
        self.reminderTime = datetime.now()
        self.startDate = datetime.now()
        self.endDate = datetime.now()


class Application(QApplication):
    def __init__(self, argv):
        super().__init__(argv)
        self.setApplicationName("alarm-clock")
        self.setApplicationDisplayName("Alarm Clock")
        self.setApplicationVersion("1.2.0")

        self.debug = os.environ.get("ALARMCLOCK_DEBUG", "0") == "1"
        self.setQuitOnLastWindowClosed(self.debug)

        self.configDirectory = os.path.expanduser("~/.local/share/alarm-clock")
        self.configFile = self.configDirectory + "/config.json"

        self.trayIcon = QSystemTrayIcon(self)
        self.trayIcon.setIcon(QIcon.fromTheme("alarm-symbolic"))
        self.trayIcon.setToolTip("Alarm Clock")

        self.trayMenu = QMenu()
        self.trayIcon.activated.connect(self.openMainWindow)
        self.trayMenu.addAction("&Open").triggered.connect(self.openMainWindow)
        self.trayMenu.addAction("&Quit").triggered.connect(self.quit)
        self.trayIcon.setContextMenu(self.trayMenu)

        self.trayIcon.show()

        self.alarms: list[Alarm] = []
        self.outlookReminders: list[OutlookReminder] = []
        self.outlookToken = ""

        self.notificationProcesses: list[subprocess.Popen] = []

        self.last_tick = datetime.now()
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.tick)
        self.timer.setInterval(500)
        self.timer.start()

    def tick(self):
        current_tick = datetime.now()

        for alarm in self.alarms:
            if not alarm.enabled:
                continue
            if len(alarm.repeat) > 0 and current_tick.weekday() not in alarm.repeat:
                continue
            if alarm.time < self.last_tick.time() or alarm.time > current_tick.time():
                continue

            summary = alarm.name
            if alarm.name == "":
                summary = "Alarm"

            body = (
                "It's "
                + alarm.time.strftime("%X")
                + ", your alarm "
                + alarm.name
                + " is going off!"
            )

            notificationProcess = subprocess.Popen(
                [
                    "notify-send",
                    "--urgency=critical",
                    "--expire-time=60000",
                    "--wait",
                    "--app-name=Alarm Clock",
                    "--icon=alarm-symbolic",
                    summary,
                    body,
                ],
            )
            self.notificationProcesses.append(notificationProcess)

            if len(alarm.repeat) == 0:
                alarm.enabled = False
                mainWindow.saveConfig()

            mainWindow.reloadAlarms()

        for reminder in self.outlookReminders:
            if reminder.reminderTime < self.last_tick.astimezone(
                timezone.utc
            ) or reminder.reminderTime > current_tick.astimezone(timezone.utc):
                continue

            summary = reminder.subject

            formattedStartTime = (
                reminder.startDate.astimezone(current_tick.tzinfo)
                .time()
                .strftime("%H:%M:%S")
            )
            formattedEndTime = (
                reminder.endDate.astimezone(current_tick.tzinfo)
                .time()
                .strftime("%H:%M:%S")
            )

            body = (
                "In "
                + reminder.location
                + ", at "
                + formattedStartTime
                + " - "
                + formattedEndTime
            )

            notificationProcess = subprocess.Popen(
                [
                    "notify-send",
                    "--urgency=critical",
                    "--expire-time=60000",
                    "--wait",
                    "--app-name=Alarm Clock",
                    "--icon=alarm-symbolic",
                    summary,
                    body,
                ],
            )
            self.notificationProcesses.append(notificationProcess)

        self.last_tick = current_tick

    def openMainWindow(self):
        mainWindow.show()
        mainWindow.raise_()
        mainWindow.activateWindow()


class EditAlarmWindow(QWidget):
    def __init__(self, alarm: Alarm | None = None):
        super().__init__()

        self.alarm = alarm

        if alarm is None:
            self.setWindowTitle("Create new alarm")
            self.alarm = Alarm()
        else:
            self.setWindowTitle("Edit alarm")

        self.boxLayout = QVBoxLayout(self)

        self.timeEntry = QTimeEdit(self)

        timeFont = self.timeEntry.font()
        timeFont.setPointSize(16)
        self.timeEntry.setFont(timeFont)
        self.timeEntry.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self.timeEntry.setDisplayFormat("HH:mm:ss")
        self.timeEntry.setTime(
            QTime(self.alarm.time.hour, self.alarm.time.minute, self.alarm.time.second)
        )
        self.boxLayout.addWidget(self.timeEntry)

        self.repeatOptionsLayout = QVBoxLayout()
        self.repeatOptionsLayout.setContentsMargins(0, 0, 0, 0)
        self.repeatOptionsLayout.setSpacing(0)

        self.repeatOnce = QRadioButton("Once", self)
        self.repeatOnce.toggled.connect(self.updateRepeatState)
        self.repeatOnce.setToolTip(
            "Set to never repeat, this means that the alarm will only be run once and then disabled"
        )
        self.repeatOptionsLayout.addWidget(self.repeatOnce)

        self.repeatEveryBusinessDay = QRadioButton("Every business day", self)
        self.repeatEveryBusinessDay.toggled.connect(self.updateRepeatState)
        self.repeatEveryBusinessDay.setToolTip("Repeat Monday to Friday")
        self.repeatOptionsLayout.addWidget(self.repeatEveryBusinessDay)

        self.repeatEveryDay = QRadioButton("Every day", self)
        self.repeatEveryDay.toggled.connect(self.updateRepeatState)
        self.repeatEveryDay.setToolTip("Repeat every day in the week")
        self.repeatOptionsLayout.addWidget(self.repeatEveryDay)

        self.repeatCustom = QRadioButton("Custom repeat", self)
        self.repeatCustom.toggled.connect(self.updateRepeatState)
        self.repeatCustom.setToolTip("Set your own days you want it to repeat to")
        self.repeatOptionsLayout.addWidget(self.repeatCustom)

        self.boxLayout.addLayout(self.repeatOptionsLayout)

        self.repeatLayout = QHBoxLayout()
        self.repeatCheckBoxes = []
        for day in range(7):
            repeatCheckBox = QPushButton(calendar.day_abbr[day], self)
            repeatCheckBox.setFixedWidth(32)
            repeatCheckBox.setDisabled(True)
            effect = QGraphicsOpacityEffect(repeatCheckBox)
            effect.setEnabled(True)
            effect.setOpacity(0.4)
            repeatCheckBox.setGraphicsEffect(effect)

            # day=day is required to freeze the current day, else Python would use the last day variable
            repeatCheckBox.clicked.connect(
                lambda event, day=day: self.toggleRepeatButton(day)
            )

            self.repeatLayout.addWidget(repeatCheckBox)
            self.repeatCheckBoxes.append(repeatCheckBox)

        self.unsavedRepeatDays = []
        for day in self.alarm.repeat:
            self.toggleRepeatButton(day)

        if len(self.alarm.repeat) == 0:
            self.repeatOnce.setChecked(True)
        elif len(self.alarm.repeat) == 7:
            self.repeatEveryDay.setChecked(True)
        elif json.dumps(self.alarm.repeat) == "[0, 1, 2, 3, 4]":
            self.repeatEveryBusinessDay.setChecked(True)
        else:
            self.repeatCustom.setChecked(True)

        self.repeatOptionsLayout.addLayout(self.repeatLayout, 0)

        self.boxLayout.addWidget(QLabel("Name:", self))

        self.nameEntry = QLineEdit(self)
        self.nameEntry.setPlaceholderText("Morning Alarm")
        self.nameEntry.setText(self.alarm.name)
        self.nameEntry.returnPressed.connect(self.save)
        self.boxLayout.addWidget(self.nameEntry)

        self.buttonsLayout = QHBoxLayout()
        self.buttonsLayout.setAlignment(
            Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignBottom
        )

        self.saveButton = QPushButton(self)
        self.saveButton.clicked.connect(self.save)
        if alarm is None:
            self.saveButton.setText("Add")
        else:
            self.saveButton.setText("Save")
        self.buttonsLayout.addWidget(self.saveButton)

        self.cancelButton = QPushButton("Cancel", self)
        self.cancelButton.clicked.connect(self.close)
        self.buttonsLayout.addWidget(self.cancelButton)

        self.boxLayout.addLayout(self.buttonsLayout, 1)

    def toggleRepeatButton(self, day):
        if day in self.unsavedRepeatDays:
            self.unsavedRepeatDays.remove(day)
            self.repeatCheckBoxes[day].setStyleSheet("")
        else:
            self.unsavedRepeatDays.append(day)
            self.unsavedRepeatDays.sort()
            self.repeatCheckBoxes[day].setStyleSheet(
                "background: " + self.palette().base().color().name() + ";"
            )

    def updateRepeatState(self, check):
        if not check:
            return

        if self.repeatOnce.isChecked():
            for day in list(self.unsavedRepeatDays):
                self.toggleRepeatButton(day)
        elif self.repeatEveryDay.isChecked():
            for day in list(self.unsavedRepeatDays):
                self.toggleRepeatButton(day)

            for day in range(7):
                self.toggleRepeatButton(day)
        elif self.repeatEveryBusinessDay.isChecked():
            for day in list(self.unsavedRepeatDays):
                self.toggleRepeatButton(day)

            for day in range(5):
                self.toggleRepeatButton(day)

        isCustom = self.repeatCustom.isChecked()
        for checkbox in self.repeatCheckBoxes:
            checkbox.setDisabled(not isCustom)
            checkbox.graphicsEffect().setEnabled(not isCustom)

    def save(self):
        self.alarm.repeat = list(self.unsavedRepeatDays)

        newTime = self.timeEntry.time()
        self.alarm.time = time(newTime.hour(), newTime.minute(), newTime.second())

        self.alarm.name = self.nameEntry.text().strip()

        if self.alarm not in app.alarms:
            app.alarms.append(self.alarm)

        mainWindow.reloadAlarms()
        mainWindow.saveConfig()

        self.close()


class AlarmEntryWidget(QFrame):
    def __init__(self, parent):
        super().__init__(parent)

        self.alarm: Alarm = None

        self.setObjectName("alarmEntryWidget")

        self.setStyleSheet(
            "#alarmEntryWidget {background: "
            + self.palette().alternateBase().color().name()
            + "; border-radius: 4px; border: 1px solid "
            + self.palette().mid().color().name()
            + ";}"
        )

        self.boxLayout = QHBoxLayout(self)
        self.boxLayout.setContentsMargins(16, 8, 16, 8)
        self.boxLayout.setSpacing(12)

        self.enabledCheckbox = QCheckBox(self)
        self.enabledCheckbox.toggled.connect(self.onEnabledToggle)
        self.boxLayout.addWidget(self.enabledCheckbox)

        self.leftSection = QVBoxLayout()
        self.boxLayout.addLayout(self.leftSection, 1)
        self.leftSection.setContentsMargins(0, 0, 0, 0)
        self.leftSection.setSpacing(4)

        self.titleLabel = QLabel(self)
        self.titleLabel.setFont(createCustomFont(14, 700))
        self.leftSection.addWidget(self.titleLabel)

        self.timeLabel = QLabel(self)
        self.timeLabel.setFont(createCustomFont(12))
        self.leftSection.addWidget(self.timeLabel)

        self.actionsLayout = QHBoxLayout()
        self.actionsLayout.setContentsMargins(0, 0, 0, 0)
        self.actionsLayout.setSpacing(4)

        self.editButton = QPushButton(self)
        self.editButton.setFlat(True)
        if QIcon.hasThemeIcon("document-edit-symbolic") or QIcon.hasThemeIcon(
            "document-edit"
        ):
            self.editButton.setIcon(QIcon.fromTheme("document-edit-symbolic"))
        else:
            darkMode = self.palette().alternateBase().color().red() < 128
            editIconPaths = (
                [
                    "/usr/share/icons/Papirus-Dark/symbolic/actions/document-edit-symbolic.svg",
                    "/usr/share/icons/breeze-dark/actions/symbolic/document-edit-symbolic.svg",
                    "/usr/share/icons/breeze-dark/actions/22/document-edit.svg",
                    "/usr/share/icons/breeze-dark/actions/16/document-edit.svg",
                    "/usr/share/icons/breeze-dark/actions/32/document-edit.svg",
                    "/usr/share/icons/Adwaita/symbolic/actions/document-edit-symbolic.svg",
                ]
                if darkMode
                else [
                    "/usr/share/icons/Adwaita/symbolic/actions/document-edit-symbolic.svg",
                    "/usr/share/icons/Papirus/symbolic/actions/document-edit-symbolic.svg",
                    "/usr/share/icons/breeze/actions/symbolic/document-edit-symbolic.svg",
                    "/usr/share/icons/breeze/actions/22/document-edit.svg",
                    "/usr/share/icons/breeze/actions/16/document-edit.svg",
                    "/usr/share/icons/breeze/actions/32/document-edit.svg",
                ]
            )
            for i in editIconPaths:
                if os.path.exists(i):
                    self.editButton.setIcon(QIcon(i))
                    break
        self.editButton.setIconSize(QSize(20, 20))
        self.editButton.clicked.connect(self.editAlarm)
        self.actionsLayout.addWidget(
            self.editButton,
            0,
            Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter,
        )

        self.removeButton = QPushButton(self)
        self.removeButton.setFlat(True)
        self.removeButton.setIcon(QIcon.fromTheme("edit-delete-symbolic"))
        self.removeButton.setIconSize(QSize(20, 20))
        self.removeButton.clicked.connect(self.removeAlarm)
        self.actionsLayout.addWidget(
            self.removeButton,
            0,
            Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter,
        )

        self.boxLayout.addLayout(self.actionsLayout, 0)

    def loadFromAlarm(self, alarm: Alarm):
        self.alarm = alarm

        if alarm.name == "":
            self.titleLabel.setText("Untitled alarm")
            font = self.titleLabel.font()
            font.setItalic(True)
            self.titleLabel.setFont(font)
            self.titleLabel.setDisabled(True)
        else:
            self.titleLabel.setText(alarm.name)
            font = self.titleLabel.font()
            font.setItalic(False)
            self.titleLabel.setFont(font)
            self.titleLabel.setDisabled(False)

        time_text = alarm.time.strftime("%X")

        if len(alarm.repeat) == 7:
            every_text = "every day</font>"
        elif len(alarm.repeat) == 0:
            every_text = "once</font>"
        else:
            every_text = "every</font> " + ", ".join(
                map(lambda x: calendar.day_abbr[x], alarm.repeat)
            )

        self.timeLabel.setText(
            time_text + ' <font color="' + getGrayColor() + '">' + every_text
        )

        self.enabledCheckbox.setChecked(alarm.enabled)

    def onEnabledToggle(self, newValue):
        if newValue == self.alarm.enabled:
            return

        self.alarm.enabled = newValue
        mainWindow.reloadAlarms()
        mainWindow.saveConfig()

    def editAlarm(self):
        self.editAlarmWindow = EditAlarmWindow(self.alarm)
        self.editAlarmWindow.show()

    def removeAlarm(self):
        app.alarms.remove(self.alarm)

        mainWindow.reloadAlarms()
        mainWindow.saveConfig()


class PreferencesWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.boxLayout = QVBoxLayout(self)
        self.setMaximumSize(0, 0)

        self.autostartFilepath = os.path.expanduser(
            "~/.config/autostart/alarm-clock.desktop"
        )

        self.autostartCheckBox = QCheckBox("Start automatically on boot", self)
        if os.path.exists(self.autostartFilepath):
            with open(self.autostartFilepath) as f:
                content = f.read()
                if (
                    "Hidden=true" not in content
                    and "X-GNOME-Autostart-enabled=false" not in content
                ):
                    self.autostartCheckBox.setChecked(True)
        self.autostartCheckBox.toggled.connect(self.onAutostartChange)

        self.boxLayout.addWidget(self.autostartCheckBox)

    def onAutostartChange(self, newAutostart):
        if not newAutostart:
            if os.path.exists(self.autostartFilepath):
                os.remove(self.autostartFilepath)
            return

        with open(self.autostartFilepath, "w+") as f:
            f.write(f"""[Desktop Entry]
Name=Alarm Clock
Comment=An alarm clock in your system tray
Icon=alarm-clock
Exec="{__file__}" hidden
Terminal=false
Type=Application
Categories=""")


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.resize(800, 500)

        self.toolBar = self.addToolBar("")
        self.toolBar.setObjectName("toolBar")
        self.toolBar.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextBesideIcon)
        self.toolBar.addAction(
            QIcon.fromTheme("list-add-symbolic"), "Add new alarm…"
        ).triggered.connect(self.addNewAlarm)
        self.toolBar.addAction(
            QIcon.fromTheme("preferences-system-symbolic"), "Preferences"
        ).triggered.connect(self.openPreferences)

        self.outlookAction = self.toolBar.addAction(
            QIcon.fromTheme("mail-client"), "Outlook"
        )
        self.outlookMenu = QMenu()
        self.outlookMenu.addAction("Synchronize...").triggered.connect(
            self.synchronizeOutlook
        )
        self.reminderCountAction = self.outlookMenu.addAction("")
        self.reminderCountAction.setDisabled(True)
        self.outlookAction.setMenu(self.outlookMenu)
        self.outlookAction.triggered.connect(self.openOutlookMenu)

        self.toolBar.setMovable(False)
        self.toolBar.setStyleSheet("#toolBar {border: none;}")

        self.mainWidget = QScrollArea(self)
        self.mainWidget.setWidgetResizable(True)
        self.setCentralWidget(self.mainWidget)

        self.mainScrollWidget = QWidget(self.mainWidget)
        self.mainWidget.setWidget(self.mainScrollWidget)

        self.boxLayout = QVBoxLayout(self.mainScrollWidget)
        self.boxLayout.setAlignment(Qt.AlignmentFlag.AlignTop)

        self.noAlarmsWidget = QWidget(self.mainScrollWidget)
        self.noAlarmsLayout = QVBoxLayout(self.noAlarmsWidget)
        self.noAlarmsLayout.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self.noAlarmsIcon = QLabel(self.noAlarmsWidget)
        self.noAlarmsIcon.setPixmap(
            QIcon.fromTheme("dialog-question-symbolic").pixmap(32, 32)
        )
        self.noAlarmsIcon.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.noAlarmsLayout.addWidget(self.noAlarmsIcon)

        self.noAlarmsText = QLabel(
            'You don\'t have any alarms yet.\nClick on "Add new alarm..." to add a new one.',
            self.noAlarmsWidget,
        )
        self.noAlarmsText.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.noAlarmsLayout.addWidget(self.noAlarmsText)

        self.noAlarmsButton = QPushButton("Add new alarm...", self.noAlarmsWidget)
        self.noAlarmsButton.setMinimumWidth(140)
        self.noAlarmsButton.clicked.connect(self.addNewAlarm)
        self.noAlarmsLayout.addWidget(
            self.noAlarmsButton, 0, Qt.AlignmentFlag.AlignCenter
        )

        self.noAlarmsWidget.hide()

        self.existingAlarmEntries = []

        self.loadConfig()

    def addNewAlarm(self):
        self.newAlarmWindow = EditAlarmWindow()
        self.newAlarmWindow.show()
        self.newAlarmWindow.move(
            int(self.x() + self.width() / 2 - self.newAlarmWindow.width() / 2),
            int(self.y() + self.height() / 2 - self.newAlarmWindow.height() / 2),
        )

    def openPreferences(self):
        self.preferencesWindow = PreferencesWindow()
        self.preferencesWindow.show()
        self.preferencesWindow.move(
            int(self.x() + self.width() / 2 - self.preferencesWindow.width() / 2),
            int(self.y() + self.height() / 2 - self.preferencesWindow.height() / 2),
        )

    def openOutlookMenu(self):
        self.outlookMenu.move(QCursor.pos())
        self.outlookMenu.show()

    def connectWithOutlook(self):
        try:
            from selenium import webdriver  # type: ignore
            from selenium.webdriver import ChromeOptions, ChromeService  # type: ignore
            from selenium.common.exceptions import NoSuchDriverException  # type: ignore
        except ImportError:
            QMessageBox.critical(
                self, "Missing libraries", "You need Selenium for Python installed."
            )
            return

        options = ChromeOptions()
        options.set_capability("goog:loggingPrefs", {"performance": "ALL"})
        options.add_argument("user-data-dir=" + app.configDirectory + "/browser")

        service = ChromeService(executable_path=shutil.which("chromedriver"))

        try:
            driver = webdriver.Chrome(options=options, service=service)
        except NoSuchDriverException:
            QMessageBox.critical(
                self,
                "Missing libraries",
                "You have Selenium for Python installed, but are missing the ChromeDriver.",
            )
            return

        driver.get("https://outlook.office.com")

        try:
            while "https://login.microsoftonline.com" not in driver.current_url:
                driver.implicitly_wait(1)

            while "https://outlook.office.com/mail" not in driver.current_url:
                driver.implicitly_wait(1)
        except Exception:
            return

        app.outlookToken = ""
        while app.outlookToken == "":
            for i in driver.get_log("performance"):
                message = json.loads(i["message"])["message"]

                auth_token = (
                    message.get("params", {})
                    .get("request", {})
                    .get("headers", {})
                    .get("authorization", "")
                )
                if auth_token == "":
                    continue

                resp = requests.post(
                    "https://outlook.office.com/owa/service.svc",
                    headers={"authorization": auth_token},
                )

                # Microsoft would respond with 404 in success
                if resp.status_code != 404:
                    continue

                app.outlookToken = auth_token
                break
            sleep(1)

        driver.close()

        self.synchronizeOutlook()

    def synchronizeOutlook(self):
        if app.outlookToken == "":
            self.connectWithOutlook()
            return

        now = datetime.now(timezone.utc)

        postData = {
            "__type": "GetRemindersJsonRequest:#Exchange",
            "Header": {
                "__type": "JsonRequestHeaders:#Exchange",
                "RequestServerVersion": "V2018_01_08",
            },
            "Body": {
                "__type": "GetRemindersRequest:#Exchange",
                "BeginTime": now.strftime("%Y-%m-%dT%H:%M:%SZ"),
                "EndTime": (now + timedelta(days=30)).strftime("%Y-%m-%dT%H:%M:%SZ"),
                "ReminderType": 1,
            },
        }

        resp = requests.post(
            "https://outlook.office.com/owa/service.svc",
            headers={
                "authorization": app.outlookToken,
                "action": "GetReminders",
                "x-owa-urlpostdata": urllib.parse.quote(json.dumps(postData)),
            },
        )

        if resp.status_code != 200:
            self.connectWithOutlook()
            return

        app.outlookReminders = []
        for rawReminder in resp.json()["Body"]["Reminders"]:
            reminderTime = datetime.fromisoformat(rawReminder["ReminderTime"])

            if reminderTime < now:
                continue

            outlookReminder = OutlookReminder()
            outlookReminder.id = rawReminder["UID"]
            outlookReminder.subject = rawReminder["Subject"]
            outlookReminder.location = rawReminder["Location"]
            outlookReminder.reminderTime = datetime.fromisoformat(
                rawReminder["ReminderTime"]
            )
            outlookReminder.startDate = datetime.fromisoformat(rawReminder["StartDate"])
            outlookReminder.endDate = datetime.fromisoformat(rawReminder["EndDate"])

            app.outlookReminders.append(outlookReminder)

        self.saveConfig()
        self.reloadAlarms()

        QMessageBox.information(
            self,
            "Outlook Reminders",
            "Successfully imported "
            + str(len(app.outlookReminders))
            + " reminders from Outlook.",
        )

    def loadConfig(self):
        if not os.path.exists(app.configFile):
            self.reloadAlarms()
            return

        with open(app.configFile) as f:
            config = json.load(f)
            for rawAlarm in config["alarms"]:
                alarm = Alarm()
                alarm.name = rawAlarm["name"]
                alarm.repeat = rawAlarm["repeat"]
                alarm.time = time.fromisoformat(rawAlarm["time"])
                alarm.enabled = rawAlarm["enabled"]

                app.alarms.append(alarm)

            for rawReminder in config.get("outlookReminders", []):
                reminder = OutlookReminder()
                reminder.id = rawReminder["id"]
                reminder.subject = rawReminder["subject"]
                reminder.location = rawReminder["location"]
                reminder.reminderTime = datetime.fromisoformat(
                    rawReminder["reminderTime"]
                )
                reminder.startDate = datetime.fromisoformat(rawReminder["startDate"])
                reminder.endDate = datetime.fromisoformat(rawReminder["endDate"])

                app.outlookReminders.append(reminder)

            app.outlookToken = config.get("outlookToken", "")

        self.reloadAlarms()

    def saveConfig(self):
        os.makedirs(app.configDirectory, exist_ok=True)

        config = {}
        config["alarms"] = []
        for alarm in app.alarms:
            rawAlarm = {}
            rawAlarm["name"] = alarm.name
            rawAlarm["repeat"] = alarm.repeat
            rawAlarm["time"] = alarm.time.isoformat()
            rawAlarm["enabled"] = alarm.enabled

            config["alarms"].append(rawAlarm)

        config["outlookReminders"] = []
        for reminder in app.outlookReminders:
            rawReminder = {}
            rawReminder["id"] = reminder.id
            rawReminder["subject"] = reminder.subject
            rawReminder["location"] = reminder.location
            rawReminder["reminderTime"] = reminder.reminderTime.isoformat()
            rawReminder["startDate"] = reminder.startDate.isoformat()
            rawReminder["endDate"] = reminder.endDate.isoformat()

            config["outlookReminders"].append(rawReminder)

        config["outlookToken"] = app.outlookToken

        with open(app.configFile, "w+") as f:
            json.dump(config, f)

    def reloadAlarms(self):
        if len(app.alarms) == 0:
            self.noAlarmsWidget.show()
            self.boxLayout.addWidget(self.noAlarmsWidget, 1)
        else:
            self.noAlarmsWidget.hide()
            self.boxLayout.removeWidget(self.noAlarmsWidget)

        for i in range(len(app.alarms) - len(self.existingAlarmEntries)):
            alarmEntry = AlarmEntryWidget(self.mainScrollWidget)
            self.boxLayout.addWidget(alarmEntry, 0)
            self.existingAlarmEntries.append(alarmEntry)

        sortedAlarms = sorted(app.alarms, key=lambda alarm: alarm.time)
        canBeNextAlarm = True
        currentTime = datetime.now()

        app.trayIcon.setToolTip("Alarm Clock")

        for i, alarmEntry in enumerate(list(self.existingAlarmEntries)):
            if i >= len(app.alarms):
                alarmEntry.deleteLater()
                self.existingAlarmEntries.pop()
                continue

            alarm = sortedAlarms[i]

            if (
                canBeNextAlarm
                and alarm.enabled
                and currentTime.weekday() in alarm.repeat
                and alarm.time > currentTime.time()
            ):
                canBeNextAlarm = False
                app.trayIcon.setToolTip(
                    "Next alarm: " + alarm.name + " at " + alarm.time.strftime("%X")
                )

            alarmEntry.loadFromAlarm(alarm)

        self.reminderCountAction.setText(str(len(app.outlookReminders)) + " reminders")

    def closeEvent(self, event):
        if app.debug:
            app.quit()

        return super().closeEvent(event)


if __name__ == "__main__":
    os.environ["QT_QPA_PLATFORMTHEME"] = ""  # Allows dark mode in Qt6
    locale.setlocale(locale.LC_ALL, "")

    app = Application(sys.argv)
    if app.palette().base().color().red() < 128:
        originalIconTheme = QIcon.themeName()
        QIcon.setThemeName("breeze-dark")
        if not QIcon.hasThemeIcon("list-add-symbolic"):
            QIcon.setThemeName(originalIconTheme)

    mainWindow = MainWindow()
    if "hidden" not in sys.argv:
        mainWindow.show()

    exitCode = app.exec()

    # Cancel open notifications with CTRL+C signal
    for i in app.notificationProcesses:
        i.send_signal(signal.SIGINT)

    sys.exit(exitCode)
