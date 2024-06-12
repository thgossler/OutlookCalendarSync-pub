<!-- SHIELDS -->
<div align="center">

[![Contributors][contributors-shield]][contributors-url]
[![Forks][forks-shield]][forks-url]
[![Stargazers][stars-shield]][stars-url]
[![Issues][issues-shield]][issues-url]

</div>

<!-- PROJECT LOGO -->
<br />
<div align="center">
  <a href="https://github.com/thgossler/OutlookCalendarSync-pub">
    <img src="images/logo.png" alt="Logo" width="80" height="80">
  </a>

  <h1 align="center">OutlookCalendarSync</h1>

  <p align="center">
    Sync your Outlook Desktop calendar selectively to your iCloud calendar.
    <br />
    <a href="https://github.com/thgossler/OutlookCalendarSync-pub/issues">Report Bug</a>
    ·
    <a href="https://github.com/thgossler/OutlookCalendarSync-pub/issues">Request Feature</a>
    ·
    <a href="https://github.com/sponsors/thgossler">Sponsor project</a>
    .
    <a href="https://www.paypal.com/donate/?hosted_button_id=JVG7PFJ8DMW7J">Donate via PayPal</a>
  </p>
</div>


## Overview

This Windows desktop program syncs your Microsoft Outlook Desktop calendar selectively to your iCloud calendar.

I worked on this some time ago and it's been doing a great job for me. Recently, I decided to share it with 
the community. I have polished it a bit for being usable by a wider audience. I hope you find it useful too. 
Be fair and consider donating if you like it.

And here's the typical use scenario for the tool, you may find your personal use case in it:

Imagine you're an employee at a company that uses Microsoft Outlook for scheduling and managing appointments. 
Your employer has strict policies and does not provide an iPhone for work, nor do they allow you to use your 
personal phone for work-related activities, nor do they allow you to connect a sync tool to your Microsoft 
Graph, nor do they allow you to install iCloud for Windows and sync the whole calendar. This can be a bit 
of a hassle, especially when you want to keep track of your work appointments outside of your office hours.

This is where the OutlookCalendarSync tool comes in handy. It's designed to sync your Outlook calendar 
selectively to your personal iCloud calendar via the Outlook Desktop application. The tool respects your privacy 
and your company's data protection concerns by only syncing the titles and locations of your work appointments, 
nothing more, directly from your local desktop to your iCloud calendar. This way, you can stay updated with 
your work schedule without compromising on your privacy or violating your company's policies.

So, even if you're away from your work computer, you can still have a glance at your upcoming work 
appointments right from your personal device, helping you manage your time more effectively. And, when you
don't want to be confronted with work appointments, just hide the corresponding iCloud calendar.

Let me know whether this fits your needs or if you have any suggestions for improvements.


## Features

- Synchronizes Outlook calendar events to an iCloud calendar (one-way, filtered)
- Default configuration settings are privacy by default and design (see data privacy policy)
- Syncs reliably into an existing iCloud calendar (empty or not) without affecting existing entries
- Flexible configuration via SyncSettings.json (auto-reload) and/or config UI
- Selectively ignores events based on blacklists of regular expressions
- Configure regular expression transformation patterns for event titles in config file
- Supports command line arguments for automation (start with /? for more information)
- Support automatic start via Windows task scheduler (or startup folder as fallback)
- Should work on all supported Windows 10+ versions (tested on 64 bit Intel/ARM, classic Outlook Desktop 2019+ with COM interface)
- Allows to reviewing the log stream when the console window is shown


## Usage tips

Many users might not be familiar with regular expressions, so here is a brief introduction to using Regular Expressions:

In the settings fields where regular expressions are supported, you can easily exclude an appointment by using specific patterns. Here is an example to illustrate:

**Example:**

Suppose you want to ensure that appointments with titles containing the word "Lunch" are not synchronized. You would add the following line to the appropriate settings field:

`Lunch`

This pattern will match any title containing "Lunch" and exclude it from syncing. 

If "Lunch" shall only match at the beginning of the title, you can use:

`^Lunch`

The `^` symbol indicates that the word "Lunch" must be at the beginning of the title to match.

"Lunch" at the end of a title would be matched by `Lunch$`.

Exactly "Lunch" as the title would be matched by `^Lunch$`.

Please note that these patterns are case-sensitive. To match both "Lunch" and "lunch," you would use e.g. or the other ways described below:

`^(L|l)unch`

**Important Aspects of .NET Regular Expressions**

- **Case Sensitivity:** By default, .NET regular expressions are case-sensitive. To perform a case-insensitive match, you can use the `(?i)` option. For example, `(?i)Lunch` will match "Lunch," "lunch," "LUNCH," etc.

- **Special Characters:** Certain characters have special meanings in regular expressions (e.g., `.`, `*`, `?`, `+`). To match these characters literally, you need to escape them with a backslash (`\`). For example, to match a period, use `\.`.

For a deeper understanding of .NET regular expressions, you can refer to a comprehensive [quick reference guide](https://download.microsoft.com/download/D/2/4/D240EBF6-A9BA-4E4F-A63F-AEB6DA0B921C/Regular%20expressions%20quick%20reference.pdf) or search online for resources using terms like "Learn .NET regular expressions."

By familiarizing yourself with these basics and critical aspects, you can effectively use regular expressions to manage your appointments and other tasks.


## Roadmap

- [ ] Ask to create new iCloud calendar if not exist
- [ ] Option to exclude appointments with certain Outlook categories
- [ ] Further improvements based on comunity feedback
- [ ] Support more target calendars than iCloud (e.g. Google calendar)


## Donate

If you wish to use the tool, please consider donating an amount that reflects its value to you. You can do so either via PayPal

[![Donate via PayPal](https://www.paypalobjects.com/en_US/i/btn/btn_donate_LG.gif)](https://www.paypal.com/donate/?hosted_button_id=JVG7PFJ8DMW7J)

or via [GitHub Sponsors](https://github.com/sponsors/thgossler).


<!-- MARKDOWN LINKS & IMAGES (https://www.markdownguide.org/basic-syntax/#reference-style-links) -->
[contributors-shield]: https://img.shields.io/github/contributors/thgossler/OutlookCalendarSync-pub.svg
[contributors-url]: https://github.com/thgossler/OutlookCalendarSync-pub/graphs/contributors
[forks-shield]: https://img.shields.io/github/forks/thgossler/OutlookCalendarSync-pub.svg
[forks-url]: https://github.com/thgossler/OutlookCalendarSync-pub/network/members
[stars-shield]: https://img.shields.io/github/stars/thgossler/OutlookCalendarSync-pub.svg
[stars-url]: https://github.com/thgossler/OutlookCalendarSync-pub/stargazers
[issues-shield]: https://img.shields.io/github/issues/thgossler/OutlookCalendarSync-pub.svg
[issues-url]: https://github.com/thgossler/OutlookCalendarSync-pub/issues
