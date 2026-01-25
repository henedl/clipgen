# clipgen
clipgen is a Python program that uses the [gspread library](https://docs.gspread.org) and [ffmpeg](https://www.ffmpeg.org) to quickly generate video snippets based on timestamps in a Google Sheet.
clipgen is a Python program that uses the [gspread library](https://docs.gspread.org) and [ffmpeg](https://www.ffmpeg.org) to quickly generate video snippets based on timestamps in a Google Sheet.

The program was created to speed up data processing during playtests and is provided as-is, without promise of support. The target audience of this program are User Experience Researchers and UX professionals who prefer to manage their videos locally rather than in the cloud.

# How to use
## Pre-requisites
1. Install the required Python dependencies; gspread and oauth2client.
2. Install ffmpeg and ensure it is available via your PATH. Alternatively, put ffmpeg in the same directory as clipgen.py.
3. Configure your Google Authentication per [gspread's setup guide](https://docs.gspread.org/en/master/oauth2.html); clipgen requires you to have a Google Cloud project with API access, with a OAuth credentials file on your system.

## Using clipgen
- Put clipgen.py in a folder with video recordings.
- Your Google credentials.json file should either be in the working directory, or in ~/.gspread
- Launch clipgen either interactively or through command-line arguments.
- Point it to your Google Sheet and enjoy quick video clip generation based on your timestamped notes.

## About the spreadsheet
clipgen assumes that you are using a spreadsheet with a particular layout. A reference spreadsheet is [available here](#) - feel free to make a copy and use it in your studies.

# TODO-list a.k.a. remaining work
* Clean up and share the example spreadsheet.
* Timestamp cleaning/parsing doesn't handle prefix/separator characters besides + , ;
* Open the current Sheet in browser, from the command line.
* Add ability to target only one cell. Proposed syntax: "P01.11". Should also be batchable, i.e. "P01.11 + P03.11 + P03.09". Should be available directly at (current) mode select stage.
* File size output limits (e.g. if user sets 50MB max limit in app settings, all videos are compressed below that limit).
* Implement a way to select only one video to be rendered, out of several possible clips.
* Try Sheet1, Data, data, Observations, Data set, data set, and other typical worksheet names. If none of those match, just go with the zeroth index worksheet.

## Major new features a.k.a. maybe at some point
* GUI
* Airtable support
* Excel support
* CSV support
* Composite highlight videos, with clips from multiple participants.
* Title/ending cards
* Watermarking
* Cropping and time-lapsing! For example generate a time-lapse of part of the screen, such as the minimap in a strategy game.