#!/usr/bin/env python3
import os
import sys
import json
import locale
import calendar
import requests
import urllib
import shutil
import threading
import math
from time import sleep
from datetime import datetime, time, timezone, timedelta
from PyQt6.QtCore import (
    Qt,
    QTimer,
    QTime,
    QSize,
    pyqtSignal,
    pyqtSlot,
    QVariant,
    QMetaType,
    QEvent,
    QObject,
)
from PyQt6.QtGui import QIcon, QColor, QCursor
from PyQt6.QtDBus import QDBusMessage, QDBusInterface, QDBusConnection
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
    QProgressDialog,
    QSpinBox,
)


SERVICE_NAME = "me.chocolateimage.AlarmClock"


def createCustomFont(size=10, weight=400):
    font = app.font()
    font.setPointSize(size)
    font.setWeight(weight)
    return font


def getGrayColor():
    color = app.palette().text().color()
    color.setAlpha(128)
    return color.name(QColor.NameFormat.HexArgb)


def getIcon(*names):
    for name in names:
        if QIcon.hasThemeIcon(name):
            return QIcon.fromTheme(name)

    return QIcon()


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
        self._reminderTime = datetime.now()
        self.startDate = datetime.now()
        self.endDate = datetime.now()

        self.notificationId = None

    @property
    def reminderTime(self):
        if app.forcedOutlookReminderMinutes == -1:
            return self._reminderTime
        else:
            return self.startDate - timedelta(minutes=app.forcedOutlookReminderMinutes)


class AlarmClockDBus(QObject):
    @pyqtSlot()
    def show(self):
        app.openMainWindow()


class Application(QApplication):
    def __init__(self, argv):
        super().__init__(argv)
        sessionBus = QDBusConnection.sessionBus()
        if not sessionBus.registerService(SERVICE_NAME):
            print("Foregrounding existing alarm clock instance...")
            alarmClockInterface = QDBusInterface(
                SERVICE_NAME,
                "/",
                "",
                sessionBus,
            )
            message = alarmClockInterface.call("show")
            if message.type() != QDBusMessage.MessageType.ErrorMessage:
                exit(0)

        self.alarmClockDBus = AlarmClockDBus()
        sessionBus.registerObject(
            "/", self.alarmClockDBus, QDBusConnection.RegisterOption.ExportAllSlots
        )

        self.setDesktopFileName("alarm-clock")
        self.setApplicationName("alarm-clock")
        self.setApplicationDisplayName("Alarm Clock")
        self.setApplicationVersion("1.3.1")

        self.debug = os.environ.get("ALARMCLOCK_DEBUG", "0") == "1"
        self.setQuitOnLastWindowClosed(self.debug)

        self.configDirectory = os.path.expanduser("~/.local/share/alarm-clock")
        self.configFile = self.configDirectory + "/config.json"

        self.trayIcon = QSystemTrayIcon(self)
        self.trayIcon.setIcon(getIcon("alarm-symbolic", "alarm-symbolic.symbolic"))
        self.trayIcon.setToolTip("Alarm Clock")

        self.trayMenu = QMenu()
        self.trayIcon.activated.connect(self.openMainWindow)

        self.openAction = self.trayMenu.addAction(
            QIcon.fromTheme("arrow-up-symbolic"), "&Open"
        )
        self.openAction.triggered.connect(self.openMainWindow)
        self.trayMenuBottomSeparator = self.trayMenu.addSeparator()
        self.quitAction = self.trayMenu.addAction(
            QIcon.fromTheme("application-exit-symbolic"), "&Quit"
        )
        self.quitAction.triggered.connect(self.quit)

        self.trayIcon.setContextMenu(self.trayMenu)

        self.trayIcon.show()

        self.alarms: list[Alarm] = []
        self.outlookReminders: list[OutlookReminder] = []
        self.outlookToken = ""
        self.forcedOutlookReminderMinutes = -1

        self.openNotifications: list[int] = []

        self.notificationsInterface = QDBusInterface(
            "org.freedesktop.Notifications",
            "/org/freedesktop/Notifications",
            "org.freedesktop.Notifications",
            sessionBus,
        )

        sessionBus.connect(
            "org.freedesktop.Notifications",
            "/org/freedesktop/Notifications",
            "org.freedesktop.Notifications",
            "NotificationClosed",
            self.onNotificationClosed,
        )

        self.last_tick = datetime.now()
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.tick)
        self.timer.setInterval(500)
        self.timer.start()

    def showNotification(
        self,
        replacesId=None,
        summary="",
        body="",
        actions=[],
        hints={},
        expireTimeout=-1,
    ) -> int | None:
        replacesIdVariant = QVariant(0 if replacesId is None else replacesId)
        replacesIdVariant.convert(QMetaType(QMetaType.Type.UInt.value))

        actionsVariant = QVariant(actions)
        actionsVariant.convert(QMetaType(QMetaType.Type.QStringList.value))

        message = self.notificationsInterface.call(
            "Notify",
            "Alarm Clock",
            replacesIdVariant,
            "alarm-symbolic",
            summary,
            body,
            actionsVariant,
            hints,
            expireTimeout,
        )

        if message.type() == QDBusMessage.MessageType.ErrorMessage:
            print(message.errorMessage())
            return None

        notificationId = message.arguments()[0]
        return notificationId

    def closeNotification(self, notificationId):
        if notificationId not in self.openNotifications:
            return

        notificationIdVariant = QVariant(notificationId)
        notificationIdVariant.convert(QMetaType(QMetaType.Type.UInt.value))

        self.notificationsInterface.call("CloseNotification", notificationIdVariant)

    @pyqtSlot(QDBusMessage)
    def onNotificationClosed(self, message: QDBusMessage):
        notificationId = message.arguments()[0]

        if notificationId not in self.openNotifications:
            return

        self.openNotifications.remove(notificationId)

    def showOutlookReminderNotification(self, reminder: OutlookReminder):
        if (
            reminder.notificationId is not None
            and reminder.notificationId not in self.openNotifications
        ):
            return

        current_time = datetime.now()

        summary = reminder.subject

        formattedStartTime = (
            reminder.startDate.astimezone(current_time.tzinfo)
            .time()
            .strftime("%H:%M:%S")
        )
        formattedEndTime = (
            reminder.endDate.astimezone(current_time.tzinfo).time().strftime("%H:%M:%S")
        )

        in_minutes = math.ceil(
            (reminder.startDate - current_time.astimezone(timezone.utc)).total_seconds()
            / 60
        )

        if in_minutes == 0:
            in_minutes_text = "Now"
        elif in_minutes == 1:
            in_minutes_text = "In less than a minute"
        else:
            in_minutes_text = "In " + str(in_minutes) + " minutes"

        body = (
            in_minutes_text
            + "\n"
            + reminder.location
            + ", at "
            + formattedStartTime
            + " - "
            + formattedEndTime
        )

        newNotificationId = self.showNotification(
            replacesId=reminder.notificationId,
            summary=summary,
            body=body,
            hints={"urgency": 2},
            expireTimeout=0,
        )

        if newNotificationId is None:
            return

        if reminder.notificationId is None:
            reminder.notificationId = newNotificationId

            self.openNotifications.append(reminder.notificationId)

    def tick(self):
        current_tick = datetime.now()
        current_tick_utc = current_tick.astimezone(timezone.utc)

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

            notificationId = self.showNotification(
                summary=summary,
                body=body,
                hints={"urgency": 2},
                expireTimeout=0,
            )

            if notificationId is not None:
                self.openNotifications.append(notificationId)

            if len(alarm.repeat) == 0:
                alarm.enabled = False
                mainWindow.saveConfig()

            mainWindow.reloadAlarms()

        for reminder in self.outlookReminders:
            if reminder.reminderTime > current_tick_utc:
                continue

            if (
                reminder.startDate > self.last_tick.astimezone(timezone.utc)
                and reminder.startDate < current_tick_utc
            ):
                self.showOutlookReminderNotification(reminder)
                continue

            if reminder.startDate < current_tick_utc:
                continue

            if reminder.notificationId is not None:
                next_minute = current_tick.replace(
                    second=reminder.reminderTime.second,
                    microsecond=reminder.reminderTime.microsecond,
                )

                if next_minute < self.last_tick or next_minute > current_tick:
                    continue

            self.showOutlookReminderNotification(reminder)

        self.last_tick = current_tick

    def openMainWindow(self):
        mainWindow.show()
        mainWindow.raise_()
        mainWindow.activateWindow()


class EditAlarmWindow(QWidget):
    def __init__(self, alarm: Alarm | None = None):
        super().__init__()

        self.alarm = alarm

        self.setWindowIcon(QIcon.fromTheme("alarm-symbolic"))
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

        self.timeEntry.installEventFilter(self)
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

    def eventFilter(self, obj, event):
        if obj == self.timeEntry:
            if event.type() == QEvent.Type.KeyPress:
                if event.key() == Qt.Key.Key_Return:
                    self.save()
                    return True

        return super().eventFilter(obj, event)

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

        self.setWindowIcon(QIcon.fromTheme("settings-configure-symbolic"))
        self.setWindowTitle("Preferences")

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

        rowLayout = QHBoxLayout()
        rowLayout.setContentsMargins(0, 0, 0, 0)
        label = QLabel(self)
        label.setText("Override Outlook reminder: ")
        rowLayout.addWidget(label)

        self.forcedReminderTime = QSpinBox(self)
        self.forcedReminderTime.setSuffix(" minutes")
        self.forcedReminderTime.setMinimum(-1)
        self.forcedReminderTime.setMaximum(10000)
        self.forcedReminderTime.setSpecialValueText("Default")
        self.forcedReminderTime.setValue(app.forcedOutlookReminderMinutes)
        self.forcedReminderTime.editingFinished.connect(self.onForcedReminderTimeChange)
        rowLayout.addWidget(self.forcedReminderTime)

        self.boxLayout.addLayout(rowLayout)

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

    def onForcedReminderTimeChange(self):
        app.forcedOutlookReminderMinutes = self.forcedReminderTime.value()
        mainWindow.saveConfig()


class MainWindow(QMainWindow):
    progressDialogSetValue = pyqtSignal(int)
    progressDialogSetLabel = pyqtSignal(str)
    progressDialogClose = pyqtSignal()
    openMessageBox = pyqtSignal(str, str, str)
    synchronizeOutlookFinished = pyqtSignal()

    def __init__(self):
        super().__init__()
        self.resize(800, 500)

        self.setWindowIcon(QIcon.fromTheme("alarm-symbolic"))

        self.toolBar = self.addToolBar("Toolbar")
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
        self.synchronizeAction = self.outlookMenu.addAction("Synchronize...")
        self.synchronizeAction.setIcon(QIcon.fromTheme("mail-download-later-symbolic"))
        self.synchronizeAction.triggered.connect(self.synchronizeOutlook)
        self.reminderCountAction = self.outlookMenu.addAction("")
        self.reminderCountAction.setDisabled(True)
        self.outlookAction.setMenu(self.outlookMenu)
        self.outlookAction.triggered.connect(self.openOutlookMenu)
        app.trayMenu.insertAction(app.trayMenuBottomSeparator, self.outlookAction)

        self.toolBar.setMovable(False)
        self.toolBar.setStyleSheet("#toolBar {border: none;}")
        self.toolBar.setContextMenuPolicy(Qt.ContextMenuPolicy.PreventContextMenu)

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

        self.progressDialog: QProgressDialog = None
        self.progressDialogSetValue.connect(
            lambda value: (
                self.progressDialog.setRange(0, 0 if value == -1 else 100),
                self.progressDialog.setValue(value),
            )
        )
        self.progressDialogSetLabel.connect(
            lambda text: self.progressDialog.setLabelText(text)
        )
        self.progressDialogClose.connect(lambda: self.progressDialog.close())
        self.openMessageBox.connect(self.openMessageBoxFunction)
        self.synchronizeOutlookFinished.connect(self.synchronizeOutlookFinishedFunction)

        self.loadConfig()

    def openMessageBoxFunction(self, type, title, text):
        getattr(QMessageBox, type)(self, title, text)

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
        self.progressDialogSetValue.emit(0)
        self.progressDialogSetLabel.emit("Connecting with Outlook...")
        try:
            from selenium import webdriver  # type: ignore
            from selenium.webdriver import ChromeOptions  # type: ignore
            from selenium.webdriver.chrome.service import Service as ChromeService  # type: ignore
            from selenium.common.exceptions import NoSuchDriverException  # type: ignore
        except ImportError:
            self.progressDialogClose.emit()
            self.openMessageBox.emit(
                "critical",
                "Missing libraries",
                "You need Selenium for Python installed.",
            )
            return

        self.progressDialogSetValue.emit(20)

        options = ChromeOptions()
        options.set_capability("goog:loggingPrefs", {"performance": "ALL"})
        options.add_argument("user-data-dir=" + app.configDirectory + "/browser")

        service = ChromeService(executable_path=shutil.which("chromedriver"))

        try:
            driver = webdriver.Chrome(options=options, service=service)
        except NoSuchDriverException:
            self.progressDialogClose.emit()
            self.openMessageBox.emit(
                "critical",
                "Missing libraries",
                "You have Selenium for Python installed, but are missing the ChromeDriver.",
            )
            return

        self.progressDialogSetValue.emit(40)

        driver.get("https://outlook.office.com")

        try:
            while "https://login.microsoftonline.com" not in driver.current_url:
                driver.implicitly_wait(1)

            self.progressDialogSetValue.emit(60)
            self.progressDialogSetLabel.emit("Please enter your login credentials")

            while "https://outlook.office.com/mail" not in driver.current_url:
                driver.implicitly_wait(1)

            self.progressDialogSetLabel.emit("Connecting with Outlook...")
            self.progressDialogSetValue.emit(80)
        except Exception:
            self.progressDialogClose.emit()
            self.openMessageBox.emit(
                "information",
                "Cancelled",
                "Login has been cancelled. Not synchronizing.",
            )
            return

        self.progressDialogSetValue.emit(90)
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
        self.progressDialogSetValue.emit(95)

        self.synchronizeOutlookBlocking()

    def synchronizeOutlook(self):
        self.progressDialog = QProgressDialog("Initializing...", None, 0, 100, self)
        self.progressDialog.show()
        threading.Thread(target=self.synchronizeOutlookBlocking).start()

    def synchronizeOutlookBlocking(self):
        if app.outlookToken == "":
            self.connectWithOutlook()
            return

        self.progressDialogSetValue.emit(-1)
        self.progressDialogSetLabel.emit("Downloading reminders...")

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
            outlookReminder._reminderTime = datetime.fromisoformat(
                rawReminder["ReminderTime"]
            )
            outlookReminder.startDate = datetime.fromisoformat(rawReminder["StartDate"])
            outlookReminder.endDate = datetime.fromisoformat(rawReminder["EndDate"])

            app.outlookReminders.append(outlookReminder)

        self.synchronizeOutlookFinished.emit()

    def synchronizeOutlookFinishedFunction(self):
        self.saveConfig()
        self.reloadAlarms()

        self.progressDialog.close()
        self.openMessageBox.emit(
            "information",
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
                reminder._reminderTime = datetime.fromisoformat(
                    rawReminder["reminderTime"]
                )
                reminder.startDate = datetime.fromisoformat(rawReminder["startDate"])
                reminder.endDate = datetime.fromisoformat(rawReminder["endDate"])

                app.outlookReminders.append(reminder)

            app.outlookToken = config.get("outlookToken", "")
            app.forcedOutlookReminderMinutes = config.get(
                "forcedOutlookReminderMinutes", -1
            )

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
            rawReminder["reminderTime"] = reminder._reminderTime.isoformat()
            rawReminder["startDate"] = reminder.startDate.isoformat()
            rawReminder["endDate"] = reminder.endDate.isoformat()

            config["outlookReminders"].append(rawReminder)

        config["outlookToken"] = app.outlookToken
        config["forcedOutlookReminderMinutes"] = app.forcedOutlookReminderMinutes

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

    originalIconTheme = QIcon.themeName()

    if not QIcon.hasThemeIcon("alarm-symbolic") and not QIcon.hasThemeIcon(
        "alarm-symbolic.symbolic"
    ):
        QIcon.setThemeName("Adwaita")

    app = Application(sys.argv)
    if app.palette().base().color().red() < 128:
        QIcon.setThemeName("breeze-dark")
        if not QIcon.hasThemeIcon("list-add-symbolic"):
            QIcon.setThemeName(originalIconTheme)

    mainWindow = MainWindow()
    if "hidden" not in sys.argv:
        mainWindow.show()

    exitCode = app.exec()

    # Cancel open notifications with D-Bus call
    for i in app.openNotifications:
        app.closeNotification(i)

    sys.exit(exitCode)
