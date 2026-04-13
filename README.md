# ninjav_xml_writer
reads FCPXML marker ratings and writes Premiere-compatible XMP markers into copied media files

app_realtime_48k writes labels with timebases calculated on the 48KHz audio sampling, according to the Adobe Premiere xmp file rules, while app_native_timebase does not do this calculation, insert labels with the original timebase.

exiftool.exe was used to write xmp metadata to media. often compiled together.
