# clipgen
clipgen is a Python script that uses the [GSpread library](https://docs.gspread.org) and [ffmpeg](https://www.ffmpeg.org) to quickly generate video snippets based on timestamps in a Google Sheet.

The script was created to speed up data processing during playtests and is provided as-is, without promise of support.

# TODO a.k.a. upcoming features
## Quality of life:
* Timestamp cleaning/parsing doesn't handle prefix/separator characters besides + , ;
* Open the current Sheet in Chrome from the command line
* Add ability to target only one cell. Proposed syntax: "P01.11". Should also be batchable, i.e. "P01.11 + P03.11 + P03.09". Should be available directly at (current) mode select stage.
* File size output limits (e.g. if user sets 50MB max limit in app settings, all videos are compressed below that limit)

## Programming stuff:
* Try Sheet1, Data, data, Observations, Data set, data set, and other typical worksheet names. If none of those match, just go with the zeroth index worksheet.
* Logging of which timestamps are discarded
* Better error messages
* Rename "generate"-methods to more clearly indicate that they return timestamps to clip (for generate_list(), this method should have a completely different name)
* Rename "dumped"-methods once all timestamps are generated from a dumped sheet instead of a live sheet
* Rename "users" to "participants" (variables, method names)

## Batch improvements:
* Implement a way to select only one video to be rendered, out of several possible clips

## Major new features a.k.a. maybe at some point:
* GUI
* Airtable support
* Excel support
* CSV support
* Composite highlight videos, with clips from multiple participants
* Title/ending cards
* Watermarking
* Cropping and time-lapsing! For example generate a time-lapse of part of the screen, such as the minimap in a strategy game.