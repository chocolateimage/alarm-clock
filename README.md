# Alarm Clock

A simple to use alarm clock for Linux with Outlook reminder integration.

## Installation

- **Arch Linux/AUR:** Use the [`alarm-clock`](https://aur.archlinux.org/packages/alarm-clock) AUR package.
- **Debian/Ubuntu based:** Add the [PlayLook Debian repository](https://packages.playlook.de/deb/), then install using `sudo apt install alarm-clock`.
- **Fedora:** Enable the [COPR repo](https://copr.fedorainfracloud.org/coprs/chocolateimage/alarm-clock/) and install the package as noted with the instructions there.

## Development

Requirements:

```sh
sudo apt install python3 python3-pyqt6 python3-requests
```

For developing, run `ALARMCLOCK_DEBUG=1 ./alarm-clock.py`.
