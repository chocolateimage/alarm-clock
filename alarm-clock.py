#!/usr/bin/env python3
import os
import sys
import json
import locale
import calendar
import subprocess
import signal
from datetime import datetime, time
from PyQt6.QtCore import Qt, QTimer, QTime
from PyQt6.QtGui import QIcon, QColor
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
        self.repeat = [0, 1, 2, 3, 4, 5, 6]
        self.time = time(0, 0, 0)
        self.enabled = True


class Application(QApplication):
    def __init__(self, argv):
        super().__init__(argv)
        self.setApplicationName("alarm-clock")
        self.setApplicationDisplayName("Alarm Clock")
        self.setQuitOnLastWindowClosed(False)

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
            if current_tick.weekday() not in alarm.repeat:
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

        self.boxLayout.addWidget(QLabel("Time:", self))

        self.timeEntry = QTimeEdit(self)
        self.timeEntry.setDisplayFormat("HH:mm:ss")
        self.timeEntry.setTime(
            QTime(self.alarm.time.hour, self.alarm.time.minute, self.alarm.time.second)
        )
        self.boxLayout.addWidget(self.timeEntry)

        self.boxLayout.addWidget(QLabel("Repeat:", self))

        self.repeatLayout = QHBoxLayout()
        self.repeatCheckBoxes = []
        for day in range(7):
            repeatCheckBox = QCheckBox(calendar.day_abbr[day], self)

            if day in self.alarm.repeat:
                repeatCheckBox.setChecked(True)

            self.repeatLayout.addWidget(repeatCheckBox)
            self.repeatCheckBoxes.append(repeatCheckBox)

        self.boxLayout.addLayout(self.repeatLayout, 0)

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

    def save(self):
        self.alarm.repeat.clear()
        for day, checkBox in enumerate(self.repeatCheckBoxes):
            if checkBox.isChecked():
                self.alarm.repeat.append(day)

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

        self.actionsLayout = QVBoxLayout()
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
        self.editButton.clicked.connect(self.editAlarm)
        self.actionsLayout.addWidget(
            self.editButton,
            0,
            Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignTop,
        )

        self.removeButton = QPushButton(self)
        self.removeButton.setFlat(True)
        self.removeButton.setIcon(QIcon.fromTheme("edit-delete-symbolic"))
        self.removeButton.clicked.connect(self.removeAlarm)
        self.actionsLayout.addWidget(
            self.removeButton,
            0,
            Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignTop,
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
        every_text = ", ".join(map(lambda x: calendar.day_abbr[x], alarm.repeat))
        if len(alarm.repeat) == 7:
            every_text = '<font color="' + getGrayColor() + '">day</font>'

        self.timeLabel.setText(
            time_text
            + ' <font color="'
            + getGrayColor()
            + '">every</font> '
            + every_text
        )

        self.enabledCheckbox.setChecked(alarm.enabled)

    def onEnabledToggle(self, newValue):
        if newValue == self.alarm.enabled:
            return

        self.alarm.enabled = newValue
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

    def openPreferences(self):
        self.preferencesWindow = PreferencesWindow()
        self.preferencesWindow.show()

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

        for i, alarmEntry in enumerate(list(self.existingAlarmEntries)):
            if i >= len(app.alarms):
                alarmEntry.deleteLater()
                self.existingAlarmEntries.pop()
                continue

            alarmEntry.loadFromAlarm(sortedAlarms[i])


if __name__ == "__main__":
    os.environ["QT_QPA_PLATFORMTHEME"] = "gtk3"  # Allows dark mode in Qt6
    locale.setlocale(locale.LC_ALL, "")

    app = Application(sys.argv)

    mainWindow = MainWindow()
    if "hidden" not in sys.argv:
        mainWindow.show()

    exitCode = app.exec()

    # Cancel open notifications with CTRL+C signal
    for i in app.notificationProcesses:
        i.send_signal(signal.SIGINT)

    sys.exit(exitCode)
