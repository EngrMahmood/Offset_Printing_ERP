import logging

logger = logging.getLogger(__name__)

# Planning signal handlers are intentionally disabled for SKU approval sync.
# This keeps workflow transitions explicit and avoids hidden automation
# between SKU master data approval and Planning job creation.
