"""
Dual-mode translator: offline (argostranslate) + online (OpenAI API).
Offline is prioritized; online is a fallback when models aren't available.
"""

import os

ARGOS_LANG_MAP = {
    "en": "en",
    "cn": "zh",
    "jp": "ja",
    "kr": "ko",
    "vi": "vi",
    "th": "th",
}

ARGOS_LANG_REVERSE = {v: k for k, v in ARGOS_LANG_MAP.items()}


API_PROVIDERS = {
    "OpenAI": {
        "base_url": None,
        "model": "gpt-4o-mini",
    },
    "DeepSeek": {
        "base_url": "https://api.deepseek.com",
        "model": "deepseek-chat",
    },
    "Doubao": {
        "base_url": "https://ark.cn-beijing.volces.com/api/v3",
        "model": "doubao-seed-2.0-pro",
    },
    "Claude": {
        "base_url": "https://api.anthropic.com/v1/",
        "model": "claude-sonnet-4-20250514",
    },
}


class Translator:
    def __init__(self, mode: str = "offline", api_key: str = "",
                 provider: str = "OpenAI"):
        self.mode = mode
        self.api_key = api_key
        self.provider = provider
        self._argos_ready = False

    def _ensure_argos(self):
        if self._argos_ready:
            return
        try:
            import argostranslate.package
            import argostranslate.translate
            self._argos_ready = True
        except ImportError:
            raise RuntimeError(
                "argostranslate is not installed.\n"
                "Run: pip install argostranslate"
            )

    def is_model_available(self, src: str, tgt: str) -> bool:
        """Check if an offline translation model is available for this pair."""
        try:
            self._ensure_argos()
            import argostranslate.translate
            src_code = ARGOS_LANG_MAP.get(src, src)
            tgt_code = ARGOS_LANG_MAP.get(tgt, tgt)
            installed = argostranslate.translate.get_installed_languages()
            src_lang = next((l for l in installed if l.code == src_code), None)
            if not src_lang:
                return False
            tgt_lang = next((l for l in installed if l.code == tgt_code), None)
            if not tgt_lang:
                return False
            translation = src_lang.get_translation(tgt_lang)
            return translation is not None
        except Exception:
            return False

    def download_model(self, src: str, tgt: str):
        """Download argostranslate package for a language pair."""
        self._ensure_argos()
        import argostranslate.package
        src_code = ARGOS_LANG_MAP.get(src, src)
        tgt_code = ARGOS_LANG_MAP.get(tgt, tgt)

        argostranslate.package.update_package_index()
        available = argostranslate.package.get_available_packages()

        pkg = next(
            (p for p in available
             if p.from_code == src_code and p.to_code == tgt_code),
            None
        )
        if pkg is None:
            pkg_en_tgt = next(
                (p for p in available
                 if p.from_code == "en" and p.to_code == tgt_code),
                None
            )
            pkg_src_en = next(
                (p for p in available
                 if p.from_code == src_code and p.to_code == "en"),
                None
            )
            if pkg_src_en:
                pkg_src_en.install()
            if pkg_en_tgt:
                pkg_en_tgt.install()
            return

        pkg.install()

    def translate(self, text: str, src: str, tgt: str) -> str:
        """Translate text from src language to tgt language."""
        if not text or not text.strip():
            return ""
        if src == tgt:
            return text

        if self.mode == "offline":
            return self._translate_offline(text, src, tgt)
        else:
            return self._translate_online(text, src, tgt)

    def translate_batch(self, texts: list[str], src: str, tgt: str) -> list[str]:
        """Translate a list of texts."""
        return [self.translate(t, src, tgt) for t in texts]

    def _translate_offline(self, text: str, src: str, tgt: str) -> str:
        self._ensure_argos()
        import argostranslate.translate

        src_code = ARGOS_LANG_MAP.get(src, src)
        tgt_code = ARGOS_LANG_MAP.get(tgt, tgt)

        if not self.is_model_available(src, tgt):
            try:
                self.download_model(src, tgt)
            except Exception as e:
                print(f"[WARN] Failed to download model {src}->{tgt}: {e}")
                return text

        try:
            result = argostranslate.translate.translate(text, src_code, tgt_code)
            return result if result else text
        except Exception as e:
            print(f"[WARN] Offline translation failed: {e}")
            return text

    def _translate_online(self, text: str, src: str, tgt: str) -> str:
        if not self.api_key:
            print("[WARN] No API key provided for online translation")
            return text

        try:
            from openai import OpenAI
        except ImportError:
            raise RuntimeError(
                "openai is not installed.\n"
                "Run: pip install openai"
            )

        lang_names = {
            "en": "English", "cn": "Chinese", "jp": "Japanese",
            "kr": "Korean", "vi": "Vietnamese", "th": "Thai",
        }
        src_name = lang_names.get(src, src)
        tgt_name = lang_names.get(tgt, tgt)

        cfg = API_PROVIDERS.get(self.provider, API_PROVIDERS["OpenAI"])
        client_kwargs = {"api_key": self.api_key}
        if cfg["base_url"]:
            client_kwargs["base_url"] = cfg["base_url"]
        client = OpenAI(**client_kwargs)

        try:
            resp = client.chat.completions.create(
                model=cfg["model"],
                messages=[
                    {"role": "system", "content": (
                        f"Translate the following {src_name} text to {tgt_name}. "
                        "Return ONLY the translated text, no explanations."
                    )},
                    {"role": "user", "content": text},
                ],
                temperature=0.1,
                max_tokens=1024,
            )
            return resp.choices[0].message.content.strip()
        except Exception as e:
            print(f"[WARN] Online translation failed: {e}")
            return text
