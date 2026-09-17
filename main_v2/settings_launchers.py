"""Subprocess entry points for legacy CTk settings windows."""

from __future__ import annotations

import argparse
import sys


def main(argv: list[str] | None = None) -> int:
    """Run one legacy settings dialog in an isolated Tk process.

    Args:
        argv: Optional command-line arguments. Defaults to ``sys.argv[1:]``.

    Returns:
        Process exit code.
    """
    parser = argparse.ArgumentParser(description="Launch legacy settings dialog")
    parser.add_argument(
        "dialog",
        choices=["bbb", "rfq", "asbuild", "ds"],
        help="Settings dialog to open.",
    )
    args = parser.parse_args(argv)

    import customtkinter

    root = customtkinter.CTk()
    root.withdraw()

    if args.dialog == "bbb":
        from base.bbb_settings_gui import show_bbb_settings

        show_bbb_settings(parent=root)
    elif args.dialog == "rfq":
        from RFQ.tags_rfp_compare.rfp_tags_settings_gui import show_rfp_tags_settings

        show_rfp_tags_settings(parent=root)
    elif args.dialog == "asbuild":
        from RFQ.tags_rfp_compare.rfp_tags_settings_gui import show_rfp_tags_settings
        from RFQ.tags_rfp_compare.rfp_tags_utils import (
            load_asbuild_config,
            save_asbuild_config,
        )

        show_rfp_tags_settings(
            parent=root,
            load_fn=load_asbuild_config,
            save_fn=save_asbuild_config,
            title_suffix=" (as-build)",
        )
    else:
        from RFQ.ds_compare.ds_mto_path_settings_gui import show_ds_mto_path_settings

        show_ds_mto_path_settings(parent=root)

    root.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
