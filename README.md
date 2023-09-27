# clipgen
clipgen is a Python script that uses the GSpread library and ffmpeg to quickly generate video snippets based on timestamps in a Google Sheet.

The script was created for use by the Paradox Interactive User Research team and is provided as-is, without promise of support.

# TODO a.k.a. upcoming features
## Quality of life:
* Timestamp cleaning/parsing doesn't handle prefix/separator characters besides + , ;
* Open the current Sheet in Chrome from the command line
* Composite highlight videos, with clips from multiple participants
* Title/ending cards
* Watermarking
* Add ability to target only one cell. Proposed syntax: "P01.11". Should also be batchable, i.e. "P01.11 + P03.11 + P03.09". Should be available directly at (current) mode select stage.
* Detect folder name and match to shared project - auto-select project to work in if match
* File size output limits (e.g. if user sets 50MB max limit in app settings, all videos are compressed below that limit)
## Programming stuff:
* Convert all data processing and printing to unicode
* Accept sheets regardless of capitalization (sheets named 'data set' are accepted, but 'Data set' aren't)
* Command line arguments to run everything from a prompt instead of interactively.
* Logging of which timestamps are discarded
* Better error messages
* Refactor try statements to be more efficient
* Rename "generate"-methods to more clearly indicate that they return timestamps to clip (for generate_list(), this method should have a completely different name)
* Rename "dumped"-methods once all timestamps are generated from a dumped sheet instead of a live sheet
* Rename "users" to "participants" (variables, method names)
## Batch improvements:
* Implement a way to select only one video to be rendered, out of several possible clips
* Start using the meta fields for checking which issues are already processed and what the grouping is
* Add Tag-based clip generation (e.g. every clip that affects Bus Routes)
* Being able to select multiple non-continuous lines
 ## Major new features:
* GUI
* Airtable support
* Excel support
* CSV support
* Cropping and time-lapsing! For example generate a time-lapse of part of the screen, such as the minimap in a strategy game.