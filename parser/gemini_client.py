import os

# Current Gemini models as of Aug 2026. Older IDs such as
# gemini-1.5-flash-lite are retired for new API keys.
DEFAULT_GEMINI_MODELS = (
    "gemini-3.5-flash-lite",
    "gemini-3.5-flash",
    "gemini-2.5-flash",
)


def gemini_models():
    override = (os.getenv("GEMINI_MODEL") or "").strip()
    if override:
        rest = [m for m in DEFAULT_GEMINI_MODELS if m != override]
        return [override, *rest]
    return list(DEFAULT_GEMINI_MODELS)


def _is_model_unavailable(error):
    msg = str(error).lower()
    return (
        "404" in msg
        or "not_found" in msg
        or "no longer available" in msg
        or "not found" in msg
    )


def generate_content(client, *, config, contents):
    last_error = None
    for model in gemini_models():
        try:
            print(f"  Using Gemini model: {model}")
            return client.models.generate_content(
                model=model,
                config=config,
                contents=contents,
            )
        except Exception as e:
            if _is_model_unavailable(e):
                print(f"  Gemini model {model} unavailable, trying next...")
                last_error = e
                continue
            raise
    raise last_error
