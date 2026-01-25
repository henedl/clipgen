# clipgen
clipgen is a Python program that uses the [gspread library](https://docs.gspread.org) and [ffmpeg](https://www.ffmpeg.org) to quickly generate video snippets based on timestamps in a Google Sheet.

The program was created to speed up data processing during playtests and is provided as-is, without promise of support. The target audience of this program are User Experience Researchers and UX professionals who prefer to manage their videos locally rather than in the cloud.

# How to use
## Pre-requisites
1. Install the required Python dependencies; gspread, oauth2client, and icecream.
2. Install ffmpeg and ensure it is available via your PATH. Alternatively, put ffmpeg in the same directory as clipgen.py.
3. Configure your Google Authentication per [gspread's setup guide](https://docs.gspread.org/en/master/oauth2.html); clipgen requires you to have a Google Cloud project with API access, with a OAuth credentials file on your system.

## Starting clipgen
- Put clipgen.py in a folder with video recordings.
- Your Google credentials.json file should either be in the working directory, or in ~/.gspread
- Launch clipgen either interactively or through command-line arguments.
- Point it to your Google Sheet and enjoy quick video clip generation based on your timestamped notes.

## How to use clipgen
- Can be used interactively or non-interactively, via command line argument calls.
- Several modes of generating timestamps are supported:
-- Batch
-- Single or multiple lines
-- Ranges
-- Categories
- Sheets can be browsed interactively through the program; no need to have a web browser always open.

## About the spreadsheet
clipgen assumes that you are using a spreadsheet with a particular layout. A reference spreadsheet is [available here](#) - feel free to make a copy and use it in your studies.

Timestamps must be separated by characters ```+ , ;```
Ranges must be separated by character ```-```

# TODO-list a.k.a. remaining work
* Clean up and share the example spreadsheet.
* Make all modes available implicitly from anywhere in the program.
* File size output limits (e.g. if user sets 50MB max limit in app settings, all videos are compressed below that limit).

## Major new features a.k.a. maybe at some point
* GUI
* Airtable support
* Excel support
* CSV support
* Composite highlight videos, with clips from multiple participants.
* Title/ending cards
* Watermarking
* Cropping and time-lapsing! For example generate a time-lapse of part of the screen, such as the minimap in a strategy game.