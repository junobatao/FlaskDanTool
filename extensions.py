def str_to_bool(value, default=False):
    if value is None:
        return default
    return value.strip().lower() in ("true", "1", "yes", "on")