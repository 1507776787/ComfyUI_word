import re
import zipfile

from .text_to_word import _build_download_url, _get_unique_output_path


class TextToSRTNode:
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "字幕内容": ("STRING", {
                    "multiline": True,
                    "default": "1\n00:00:00,000 --> 00:00:03,839\nIf you don't hit a jackpot within the first 10 spins.\n\n2\n00:00:03,980 --> 00:00:05,740\nNo need to top up at all in this game."
                }),
                "文件名": ("STRING", {
                    "default": "subtitles"
                }),
            },
            "optional": {
                "下载地址前缀": ("STRING", {
                    "default": ""
                }),
                "下载方式": (["ZIP下载(推荐)", "直接SRT"], {
                    "default": "ZIP下载(推荐)"
                }),
            }
        }

    RETURN_TYPES = ("STRING", "STRING")
    RETURN_NAMES = ("文件路径", "下载链接")
    OUTPUT_NODE = True
    FUNCTION = "generate_srt"
    CATEGORY = "liujian"

    _TIMECODE_RE = re.compile(
        r"^(\d{1,2}:\d{2}:\d{2}[,.]\d{1,3})\s*-->\s*(\d{1,2}:\d{2}:\d{2}[,.]\d{1,3})$"
    )

    @classmethod
    def _normalize_timestamp(cls, timestamp):
        match = re.match(r"^(\d{1,2}):(\d{2}):(\d{2})[,.](\d{1,3})$", timestamp.strip())
        if not match:
            return timestamp.strip()
        hh, mm, ss, ms = match.groups()
        ms = (ms + "00")[:3]
        return f"{hh.zfill(2)}:{mm}:{ss},{ms}"

    @classmethod
    def _normalize_srt(cls, raw_text):
        lines = raw_text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
        cues = []
        i = 0

        while i < len(lines):
            current = lines[i].strip()
            if not current:
                i += 1
                continue

            if current.isdigit() and i + 1 < len(lines):
                maybe_timecode = lines[i + 1].strip()
                if cls._TIMECODE_RE.match(maybe_timecode):
                    i += 1
                    current = lines[i].strip()

            timecode_match = cls._TIMECODE_RE.match(current)
            if not timecode_match:
                i += 1
                continue

            start_ts, end_ts = timecode_match.groups()
            i += 1
            text_lines = []

            while i < len(lines):
                s = lines[i].strip()
                if not s:
                    i += 1
                    break
                if s.isdigit() and i + 1 < len(lines) and cls._TIMECODE_RE.match(lines[i + 1].strip()):
                    break
                if cls._TIMECODE_RE.match(s):
                    break
                text_lines.append(lines[i].strip())
                i += 1

            cues.append({
                "start": cls._normalize_timestamp(start_ts),
                "end": cls._normalize_timestamp(end_ts),
                "lines": text_lines if text_lines else [""]
            })

        if not cues:
            return raw_text.strip() + "\n"

        output_blocks = []
        for idx, cue in enumerate(cues, start=1):
            block = [
                str(idx),
                f"{cue['start']} --> {cue['end']}",
                *cue["lines"]
            ]
            output_blocks.append("\n".join(block))

        return "\n\n".join(output_blocks) + "\n"

    def generate_srt(self, 字幕内容, 文件名, 下载地址前缀="", 下载方式="ZIP下载(推荐)"):
        srt_text = self._normalize_srt(字幕内容 or "")
        file_path = _get_unique_output_path(文件名, ".srt")
        with open(file_path, "w", encoding="utf-8", newline="\n") as f:
            f.write(srt_text)

        if 下载方式 == "直接SRT":
            download_url = _build_download_url(file_path, 下载地址前缀)
            return {
                "ui": {"text": [f"SRT文件: {file_path}", f"下载链接: {download_url}", "下载方式: 直接SRT"]},
                "result": (file_path, download_url)
            }

        # Browser usually previews plain-text SRT on /view.
        # Wrap it into a zip so clicking the link triggers download directly.
        zip_path = _get_unique_output_path(文件名, ".zip")
        with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            zf.write(file_path, arcname=file_path.rsplit("\\", 1)[-1])

        download_url = _build_download_url(zip_path, 下载地址前缀)
        return {
            "ui": {"text": [f"SRT文件: {file_path}", f"下载包: {zip_path}", f"下载链接: {download_url}", "下载方式: ZIP下载"]},
            "result": (file_path, download_url)
        }

    @classmethod
    def IS_CHANGED(cls, 字幕内容, 文件名, **kwargs):
        return hash((字幕内容 or "") + (文件名 or ""))


NODE_CLASS_MAPPINGS = {
    "TextToSRT": TextToSRTNode,
}

NODE_DISPLAY_NAME_MAPPINGS = {
    "TextToSRT": "字幕文本生成SRT",
}
