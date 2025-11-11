from infra.config_loader import load_config

# Mirror allowed extensions from central config
ALLOWED_EXTENSIONS = load_config().get("target_extensions", [".docx", ".pptx", ".pdf", ".xlsx"])
