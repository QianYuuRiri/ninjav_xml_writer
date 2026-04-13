import os
import sys
import uuid
import shutil
import tempfile
import threading
import queue
import subprocess
from dataclasses import dataclass
from fractions import Fraction
import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext


SAMPLE_RATE = 48000
COLOR_RED = 4281740498
COLOR_BLUE = 4294741314


@dataclass
class Marker:
    start_samples: int
    duration_samples: int
    name: str
    color: int  # 0=default(green), COLOR_RED, COLOR_BLUE


@dataclass
class MediaMarkers:
    markers: list


def _resource_path(relative_name: str) -> str:
    base_dir = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_dir, relative_name)


def _parse_time_to_fraction(value: str) -> Fraction:
    if not value:
        return Fraction(0)
    text = value.strip()
    if text.endswith("s"):
        text = text[:-1]
    if not text:
        return Fraction(0)
    if "/" in text:
        num, den = text.split("/", 1)
        return Fraction(int(num), int(den))
    return Fraction(text)


def _round_fraction_to_int(value: Fraction) -> int:
    n = value.numerator
    d = value.denominator
    if n >= 0:
        return (2 * n + d) // (2 * d)
    return -((-2 * n + d) // (2 * d))


def _seconds_to_samples(seconds: Fraction) -> int:
    return _round_fraction_to_int(seconds * SAMPLE_RATE)


def _rating_name(value: str) -> str:
    if not value:
        return ""
    return value.strip()


def _get_marker_color(value: str) -> int:
    if not value:
        return COLOR_BLUE
    lowered = value.strip().lower()
    if lowered == "favorite":
        return 0
    if lowered == "reject":
        return COLOR_RED
    return COLOR_BLUE


def _fraction_to_fps(frame_duration: Fraction) -> str:
    if frame_duration == 0:
        return "unknown"
    fps = Fraction(1, 1) / frame_duration
    if fps.denominator == 1:
        return str(fps.numerator)
    return f"{fps.numerator}/{fps.denominator}"


def parse_fcpxml(fcpxml_path: str, logger=None):
    tree = ET.parse(fcpxml_path)
    root = tree.getroot()

    formats = {}
    for fmt in root.findall(".//{*}format"):
        fmt_id = fmt.get("id")
        frame_duration = _parse_time_to_fraction(fmt.get("frameDuration"))
        if fmt_id:
            formats[fmt_id] = frame_duration

    assets = {}
    for asset in root.findall(".//{*}asset"):
        asset_id = asset.get("id")
        assets[asset_id] = {
            "src": asset.get("src", ""),
            "name": asset.get("name", ""),
            "format_id": asset.get("format"),
            "start": asset.get("start"),
        }

    markers_by_media = {}
    clip_reports = []

    for clip in root.findall(".//{*}clip"):
        clip_name = clip.get("name", "")
        clip_format_id = clip.get("format")
        tc_format = (clip.get("tcFormat") or "").upper() or "UNKNOWN"

        video = clip.find(".//{*}video")
        asset_ref = video.get("ref") if video is not None else None

        asset = assets.get(asset_ref)
        if asset is None and clip_name:
            for asset_item in assets.values():
                if asset_item.get("name") == clip_name:
                    asset = asset_item
                    break

        if asset is None:
            if logger:
                logger(f"[WARN] Asset not found for clip '{clip_name}'.")
            continue

        src = asset.get("src", "")
        media_name = os.path.basename(src).strip()
        if not media_name:
            if logger:
                logger(f"[WARN] Empty media src for clip '{clip_name}'.")
            continue

        format_id = asset.get("format_id") or clip_format_id
        frame_duration = formats.get(format_id, Fraction(0))
        fps_text = _fraction_to_fps(frame_duration)
        clip_reports.append((clip_name, media_name, tc_format, fps_text))

        media_key = media_name.lower()
        bucket = markers_by_media.get(media_key)
        if bucket is None:
            bucket = MediaMarkers(markers=[])
            markers_by_media[media_key] = bucket

        asset_start_seconds = _parse_time_to_fraction(asset.get("start"))

        for rating in clip.findall(".//{*}rating"):
            name = _rating_name(rating.get("value", ""))
            color = _get_marker_color(rating.get("value", ""))

            rating_start_seconds = _parse_time_to_fraction(rating.get("start"))
            rating_duration_seconds = _parse_time_to_fraction(rating.get("duration"))

            start_offset_seconds = rating_start_seconds - asset_start_seconds
            if start_offset_seconds < 0:
                start_offset_seconds = Fraction(0)

            start_samples = _seconds_to_samples(start_offset_seconds)
            duration_samples = _seconds_to_samples(rating_duration_seconds)

            bucket.markers.append(Marker(start_samples, duration_samples, name, color))

    return markers_by_media, clip_reports


def build_premiere_xmp(markers: list) -> str:
    lines = [
        "<?xpacket begin=\"\ufeff\" id=\"W5M0MpCehiHzreSzNTczkc9d\"?>",
        "<x:xmpmeta xmlns:x=\"adobe:ns:meta/\">",
        "  <rdf:RDF xmlns:rdf=\"http://www.w3.org/1999/02/22-rdf-syntax-ns#\">",
        "    <rdf:Description rdf:about=\"\" xmlns:xmpDM=\"http://ns.adobe.com/xmp/1.0/DynamicMedia/\">",
        "      <xmpDM:Tracks>",
        "        <rdf:Bag>",
        "          <rdf:li rdf:parseType=\"Resource\">",
        "            <xmpDM:trackName>Comment</xmpDM:trackName>",
        "            <xmpDM:trackType>Comment</xmpDM:trackType>",
        f"            <xmpDM:frameRate>f{SAMPLE_RATE}</xmpDM:frameRate>",
        "            <xmpDM:markers>",
        "              <rdf:Seq>",
    ]

    for marker in markers:
        guid = str(uuid.uuid4())
        color_guid = str(uuid.uuid4())

        lines.append("                <rdf:li rdf:parseType=\"Resource\">")
        lines.append(f"                  <xmpDM:startTime>{marker.start_samples}</xmpDM:startTime>")

        if marker.duration_samples > 0:
            lines.append(f"                  <xmpDM:duration>{marker.duration_samples}</xmpDM:duration>")

        if marker.name:
            lines.append(f"                  <xmpDM:name>{_escape_xml(marker.name)}</xmpDM:name>")

        lines.append(f"                  <xmpDM:guid>{guid}</xmpDM:guid>")

        if marker.color != 0:
            lines.extend(
                [
                    "                  <xmpDM:cuePointParams>",
                    "                    <rdf:Seq>",
                    "                      <rdf:li rdf:parseType=\"Resource\">",
                    "                        <xmpDM:key>marker_guid</xmpDM:key>",
                    f"                        <xmpDM:value>{guid}</xmpDM:value>",
                    "                      </rdf:li>",
                    "                      <rdf:li rdf:parseType=\"Resource\">",
                    f"                        <xmpDM:key>keywordExtDVAv1_{color_guid}</xmpDM:key>",
                    f'                        <xmpDM:value>{{"color":{marker.color}}}</xmpDM:value>',
                    "                      </rdf:li>",
                    "                    </rdf:Seq>",
                    "                  </xmpDM:cuePointParams>",
                ]
            )
        else:
            lines.extend(
                [
                    "                  <xmpDM:cuePointParams>",
                    "                    <rdf:Seq>",
                    "                      <rdf:li rdf:parseType=\"Resource\">",
                    "                        <xmpDM:key>marker_guid</xmpDM:key>",
                    f"                        <xmpDM:value>{guid}</xmpDM:value>",
                    "                      </rdf:li>",
                    "                    </rdf:Seq>",
                    "                  </xmpDM:cuePointParams>",
                ]
            )

        lines.append("                </rdf:li>")

    lines.extend(
        [
            "              </rdf:Seq>",
            "            </xmpDM:markers>",
            "          </rdf:li>",
            "        </rdf:Bag>",
            "      </xmpDM:Tracks>",
            "    </rdf:Description>",
            "  </rdf:RDF>",
            "</x:xmpmeta>",
            "<?xpacket end=\"w\"?>",
        ]
    )

    return "\n".join(lines) + "\n"


def _escape_xml(text: str) -> str:
    return (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace("\"", "&quot;")
        .replace("'", "&apos;")
    )


def _scan_media_files(media_dir: str) -> dict:
    file_map = {}
    for root, _, files in os.walk(media_dir):
        for name in files:
            key = name.lower()
            if key not in file_map:
                file_map[key] = os.path.join(root, name)
    return file_map


def _write_xmp_to_media(exiftool_path: str, media_path: str, xmp_content: str) -> bool:
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xmp") as temp_xmp:
        temp_xmp.write(xmp_content.encode("utf-8"))
        temp_xmp_path = temp_xmp.name

    creation_flags = 0
    if os.name == "nt":
        creation_flags = subprocess.CREATE_NO_WINDOW

    try:
        result = subprocess.run(
            [
                exiftool_path,
                "-q",
                "-q",
                "-overwrite_original",
                f"-xmp<={temp_xmp_path}",
                media_path,
            ],
            capture_output=True,
            text=True,
            creationflags=creation_flags,
        )
        return result.returncode == 0
    finally:
        try:
            os.remove(temp_xmp_path)
        except OSError:
            pass


def process_all(fcpxml_dir: str, media_dir: str, output_dir: str, exiftool_path: str, logger):
    fcpxml_files = [
        os.path.join(fcpxml_dir, name)
        for name in os.listdir(fcpxml_dir)
        if name.lower().endswith(".fcpxml")
    ]

    if not fcpxml_files:
        logger("[ERROR] No FCPXML files found.")
        return

    markers_by_media = {}
    all_clip_reports = []

    for fcpxml_path in fcpxml_files:
        fcpxml_name = os.path.basename(fcpxml_path)
        logger(f"[INFO] Parsing {fcpxml_name}...")
        parsed, clip_reports = parse_fcpxml(fcpxml_path, logger=logger)
        all_clip_reports.extend(clip_reports)
        for media_key, media_markers in parsed.items():
            bucket = markers_by_media.get(media_key)
            if bucket is None:
                markers_by_media[media_key] = media_markers
            else:
                bucket.markers.extend(media_markers.markers)

    if not markers_by_media:
        logger("[WARN] No markers found in FCPXML.")
        return

    logger("[VERIFICATION] Clip Formats:")
    for clip_name, media_name, tc_format, fps_text in all_clip_reports:
        logger(f"  - {clip_name} -> {media_name} | tc={tc_format} | fps={fps_text}")

    logger("[VERIFICATION] FCPXML Contents:")
    total_markers = sum(len(m.markers) for m in markers_by_media.values())
    logger(f"  Total media clips: {len(markers_by_media)}")
    logger(f"  Total markers: {total_markers}")
    logger("")

    for media_key, media_markers in sorted(markers_by_media.items()):
        logger(f"  - {media_key}: {len(media_markers.markers)} markers")

    media_files = _scan_media_files(media_dir)
    if not media_files:
        logger("[ERROR] No media files found in the media folder.")
        return

    logger("[VERIFICATION] Media Files:")
    found_count = 0
    missing_list = []
    for media_key in sorted(markers_by_media.keys()):
        if media_key in media_files:
            found_count += 1
            logger(f"  ✓ {media_key}")
        else:
            missing_list.append(media_key)
            logger(f"  ✗ {media_key} (NOT FOUND)")

    if missing_list:
        logger(f"[WARN] Missing {len(missing_list)} media file(s). Skipping them.")
    logger("")

    if not os.path.isfile(exiftool_path):
        logger(f"[ERROR] ExifTool not found at '{exiftool_path}'.")
        return

    os.makedirs(output_dir, exist_ok=True)

    logger("[PROCESSING] Creating output files:")
    success_count = 0
    for media_key, media_markers in markers_by_media.items():
        src_path = media_files.get(media_key)
        if not src_path:
            continue

        relative_path = os.path.relpath(src_path, media_dir)
        dest_path = os.path.join(output_dir, relative_path)
        os.makedirs(os.path.dirname(dest_path), exist_ok=True)

        try:
            shutil.copy2(src_path, dest_path)
            xmp_content = build_premiere_xmp(media_markers.markers)
            ok = _write_xmp_to_media(exiftool_path, dest_path, xmp_content)
            if ok:
                logger(f"  ✓ {os.path.basename(dest_path)}")
                success_count += 1
            else:
                logger(f"  ✗ {os.path.basename(dest_path)} (XMP write failed)")
        except Exception as exc:
            logger(f"  ✗ {os.path.basename(dest_path)} ({str(exc)})")

    logger("")
    logger(f"[DONE] Successfully processed {success_count}/{found_count} media files.")


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("FCPXML to Premiere XMP (RealTime@48k)")

        self.fcpxml_dir = tk.StringVar()
        self.media_dir = tk.StringVar()
        self.output_dir = tk.StringVar()
        self.exiftool_path = tk.StringVar()

        self.log_queue = queue.Queue()

        self._build_ui()
        self._set_defaults()
        self._poll_log_queue()

    def _build_ui(self):
        pad = {"padx": 8, "pady": 6}

        tk.Label(self.root, text="FCPXML Folder").grid(row=0, column=0, sticky="w", **pad)
        tk.Entry(self.root, textvariable=self.fcpxml_dir, width=60).grid(row=0, column=1, **pad)
        tk.Button(self.root, text="Browse", command=self._browse_fcpxml).grid(row=0, column=2, **pad)

        tk.Label(self.root, text="Media Folder").grid(row=1, column=0, sticky="w", **pad)
        tk.Entry(self.root, textvariable=self.media_dir, width=60).grid(row=1, column=1, **pad)
        tk.Button(self.root, text="Browse", command=self._browse_media).grid(row=1, column=2, **pad)

        tk.Label(self.root, text="Output Folder").grid(row=2, column=0, sticky="w", **pad)
        tk.Entry(self.root, textvariable=self.output_dir, width=60).grid(row=2, column=1, **pad)
        tk.Button(self.root, text="Browse", command=self._browse_output).grid(row=2, column=2, **pad)

        self.run_button = tk.Button(self.root, text="Run", command=self._run)
        self.run_button.grid(row=3, column=1, sticky="e", **pad)

        self.log_text = scrolledtext.ScrolledText(self.root, width=90, height=20)
        self.log_text.grid(row=4, column=0, columnspan=3, padx=8, pady=8)

    def _set_defaults(self):
        script_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
        self.fcpxml_dir.set(script_dir)
        self.media_dir.set(script_dir)
        self.output_dir.set(os.path.join(script_dir, "output"))

        candidate = _resource_path("ExifTool v12.10.exe")
        if os.path.isfile(candidate):
            self.exiftool_path.set(candidate)
        else:
            alt_candidate = _resource_path("exiftool.exe")
            if os.path.isfile(alt_candidate):
                self.exiftool_path.set(alt_candidate)

    def _browse_fcpxml(self):
        path = filedialog.askdirectory()
        if path:
            self.fcpxml_dir.set(path)

    def _browse_media(self):
        path = filedialog.askdirectory()
        if path:
            self.media_dir.set(path)

    def _browse_output(self):
        path = filedialog.askdirectory()
        if path:
            self.output_dir.set(path)

    def _run(self):
        if not self.fcpxml_dir.get() or not self.media_dir.get() or not self.output_dir.get():
            messagebox.showerror("Error", "Please select all folders.")
            return

        if not self.exiftool_path.get() or not os.path.isfile(self.exiftool_path.get()):
            messagebox.showerror("Error", "Bundled ExifTool not found.")
            return

        self.run_button.config(state="disabled")
        self.log_text.delete("1.0", tk.END)

        def worker():
            try:
                process_all(
                    self.fcpxml_dir.get(),
                    self.media_dir.get(),
                    self.output_dir.get(),
                    self.exiftool_path.get(),
                    logger=self._log,
                )
            except Exception as exc:
                self._log(f"[ERROR] {exc}")
            finally:
                self.root.after(0, lambda: self.run_button.config(state="normal"))

        threading.Thread(target=worker, daemon=True).start()

    def _log(self, message: str):
        self.log_queue.put(message)

    def _poll_log_queue(self):
        while not self.log_queue.empty():
            message = self.log_queue.get()
            self.log_text.insert(tk.END, message + "\n")
            self.log_text.see(tk.END)
        self.root.after(150, self._poll_log_queue)


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
