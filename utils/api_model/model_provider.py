"""
CAMEL ModelFactory wrapper.
Supports env var overrides: MODEL_PLATFORM, MODEL_API_KEY, MODEL_API_URL,
and sampling: MODEL_GREEDY / MODEL_TEMPERATURE / MODEL_TOP_P / MODEL_N
"""
import os
from camel.models import ModelFactory
from camel.types import ModelPlatformType
from configs.global_configs import global_configs

# provider name → (ModelPlatformType, api_key_fn, default_base_url)
_PROVIDER_MAP = {
    "aihubmix":   (ModelPlatformType.OPENAI_COMPATIBLE_MODEL,
                   lambda: global_configs.aihubmix_key,
                   "https://aihubmix.com/v1"),
    "openai":     (ModelPlatformType.OPENAI,
                   lambda: global_configs.openai_official_key, None),
    "anthropic":  (ModelPlatformType.ANTHROPIC,
                   lambda: global_configs.anthropic_official_key, None),
    "deepseek":   (ModelPlatformType.DEEPSEEK,
                   lambda: global_configs.deepseek_official_key, None),
    "openrouter": (ModelPlatformType.OPENROUTER,
                   lambda: global_configs.openrouter_key, None),
    "qwen":       (ModelPlatformType.QWEN,
                   lambda: global_configs.qwen_official_key, None),
    "gemini":     (ModelPlatformType.GEMINI,
                   lambda: global_configs.google_official_key, None),
    "openai_compatible": (ModelPlatformType.OPENAI_COMPATIBLE_MODEL,
                          lambda: os.environ.get("MODEL_API_KEY", ""), None),
}


def _sampling_config_from_env() -> dict:
    """Build CAMEL model_config_dict for decoding / sampling.

    Env vars:
      MODEL_GREEDY=1|true|yes  → temperature=0, top_p=1, n=1
      MODEL_TEMPERATURE=float  → overrides temperature
      MODEL_TOP_P=float        → overrides top_p
      MODEL_N=int              → overrides n (completions)
      MODEL_MAX_TOKENS=int     → caps each completion
      MODEL_ENABLE_THINKING=bool → toggles Qwen thinking per request
    """
    cfg: dict = {}
    greedy = os.environ.get("MODEL_GREEDY", "").strip().lower() in (
        "1", "true", "yes", "y", "on",
    )
    if greedy:
        cfg.update(temperature=0.0, top_p=1.0, n=1)

    if "MODEL_TEMPERATURE" in os.environ and os.environ["MODEL_TEMPERATURE"] != "":
        cfg["temperature"] = float(os.environ["MODEL_TEMPERATURE"])
    if "MODEL_TOP_P" in os.environ and os.environ["MODEL_TOP_P"] != "":
        cfg["top_p"] = float(os.environ["MODEL_TOP_P"])
    if "MODEL_N" in os.environ and os.environ["MODEL_N"] != "":
        cfg["n"] = int(os.environ["MODEL_N"])
    if "MODEL_MAX_TOKENS" in os.environ and os.environ["MODEL_MAX_TOKENS"] != "":
        max_tokens = int(os.environ["MODEL_MAX_TOKENS"])
        if max_tokens < 1:
            raise ValueError(
                "MODEL_MAX_TOKENS must be a positive integer, got "
                f"{max_tokens}"
            )
        cfg["max_tokens"] = max_tokens
    if (
        "MODEL_ENABLE_THINKING" in os.environ
        and os.environ["MODEL_ENABLE_THINKING"] != ""
    ):
        raw = os.environ["MODEL_ENABLE_THINKING"].strip().lower()
        if raw in ("1", "true", "yes", "y", "on"):
            enable_thinking = True
        elif raw in ("0", "false", "no", "n", "off"):
            enable_thinking = False
        else:
            raise ValueError(
                "MODEL_ENABLE_THINKING must be a boolean "
                f"(0/1, true/false), got {os.environ['MODEL_ENABLE_THINKING']!r}"
            )

        # OpenAI's extra_body is merged into the top-level request JSON by
        # CAMEL/OpenAI SDK.  SGLang consumes chat_template_kwargs directly and
        # applies Qwen's enable_thinking toggle without a server restart.
        extra_body = dict(cfg.get("extra_body") or {})
        chat_template_kwargs = dict(
            extra_body.get("chat_template_kwargs") or {}
        )
        chat_template_kwargs["enable_thinking"] = enable_thinking
        extra_body["chat_template_kwargs"] = chat_template_kwargs
        cfg["extra_body"] = extra_body
    return cfg


def build_model(model_name: str, provider: str):
    """Build a CAMEL BaseModelBackend.

    Env var overrides (take priority over provider map):
      MODEL_PLATFORM  - CAMEL ModelPlatformType name or provider key
      MODEL_API_KEY   - API key
      MODEL_API_URL   - base URL (for compatible endpoints)
      MODEL_GREEDY / MODEL_TEMPERATURE / MODEL_TOP_P / MODEL_N /
      MODEL_MAX_TOKENS / MODEL_ENABLE_THINKING - sampling, completion cap,
      and Qwen thinking toggle
    """
    # Env var overrides
    env_platform = os.environ.get("MODEL_PLATFORM")
    env_key      = os.environ.get("MODEL_API_KEY")
    env_url      = os.environ.get("MODEL_API_URL")

    if env_platform:
        # Try to resolve as provider key first, then as ModelPlatformType name
        if env_platform.lower() in _PROVIDER_MAP:
            platform, key_fn, default_url = _PROVIDER_MAP[env_platform.lower()]
        else:
            platform = ModelPlatformType[env_platform.upper()]
            key_fn = lambda: ""
            default_url = None
        api_key = env_key or key_fn()
        url = env_url or default_url
    else:
        if provider not in _PROVIDER_MAP:
            raise ValueError(
                f"Unknown provider '{provider}'. "
                f"Supported: {list(_PROVIDER_MAP.keys())}"
            )
        platform, key_fn, default_url = _PROVIDER_MAP[provider]
        api_key = env_key or key_fn()
        url = env_url or default_url

    kwargs = dict(model_platform=platform, model_type=model_name, api_key=api_key)
    if url:
        # Anthropic SDK uses base_url without /v1 suffix; others need /v1
        if platform == ModelPlatformType.ANTHROPIC:
            kwargs["url"] = url.rstrip("/").rstrip("/v1")
        else:
            kwargs["url"] = url.rstrip("/") + ("/v1" if not url.rstrip("/").endswith("/v1") else "")

    sampling = _sampling_config_from_env()
    if sampling:
        kwargs["model_config_dict"] = sampling

    return ModelFactory.create(**kwargs)
