from pyrevit import forms
from pyrevit import script

logger = script.get_logger()

forms.alert(
    "Audit placeholder. Replace this with real checks.",
    title="WWP BIM Tools",
    warn_icon=False,
)

logger.info("Audit placeholder finished.")
