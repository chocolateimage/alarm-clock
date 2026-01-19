Name: alarm-clock
Version: 1.4.0
Release: 1%{?dist}
Summary: Alarm clock

License: GPLv3
Source0: https://github.com/chocolateimage/alarm-clock/archive/refs/tags/v1.4.0.tar.gz

Requires: python3 python3-pyqt6 python3-requests
Recommends: python3-selenium chromedriver

%define debug_package %{nil}

%description
A simple alarm clock application with Outlook support

%prep
%autosetup

%install
mkdir -p %{buildroot}/usr/share/applications
mkdir -p %{buildroot}/usr/bin
install -m 644 alarm-clock.desktop %{buildroot}/usr/share/applications/alarm-clock.desktop
install -m 755 alarm-clock.py %{buildroot}/usr/bin/alarm-clock

%files
/usr/share/applications/alarm-clock.desktop
/usr/bin/alarm-clock

%changelog
* Mon Jan 19 2026 chocolateimage <chocolateimage@protonmail.com>
- Initial Fedora release
