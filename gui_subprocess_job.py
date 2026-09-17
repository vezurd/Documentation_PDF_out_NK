"""
Запуск subprocess с обновлением CTkButton: таймер, итог (успех / код / ошибка).
Все изменения виджета — только из главного потока через after().
"""

from __future__ import annotations

import threading
import time
from typing import Any, Callable, Optional, Tuple

# (returncode, exception) — либо код после wait, либо исключение до/во время wait
WorkResult = Tuple[Optional[int], Optional[BaseException]]


def _format_elapsed(seconds: float) -> str:
    if seconds < 60:
        return f"{int(seconds)} с"
    m, s = int(seconds // 60), int(seconds % 60)
    return f"{m}м {s:02d}с"


def _button_alive(button: Any) -> bool:
    try:
        return bool(button.winfo_exists())
    except Exception:
        return False


def run_subprocess_job_on_button(
    root: Any,
    button: Any,
    *,
    base_text: str,
    running_label: str,
    work_fn: Callable[[], WorkResult],
    is_running: Callable[[], bool],
    set_running: Callable[[bool], None],
    tick_ms: int = 400,
    reset_after_ms: int = 5000,
) -> None:
    """
    Старт задачи: блокирует кнопку, тикер времени, по завершении — цвет/текст и сброс через reset_after_ms.
    Если is_running() уже True — кратко показать «Уже выполняется…» (тикер это подхватит).
    """
    if is_running():
        _schedule_busy_flash(root, button)
        return

    set_running(True)

    try:
        orig_text = button.cget("text")
        orig_fg = button.cget("fg_color")
        orig_hover = button.cget("hover_color")
    except Exception:
        orig_text, orig_fg, orig_hover = base_text, None, None

    # «В работе» — нейтральный сине-серый (читается в тёмной теме CTk)
    running_fg = ("#4A6985", "#3D5A73")
    running_hover = ("#5A7A95", "#4D6A83")
    ok_fg = ("#2FA572", "#1F7A55")
    ok_hover = ("#3BB584", "#2A8A64")
    warn_fg = ("#B8860B", "#8A6508")
    warn_hover = ("#C9961B", "#9A7510")
    err_fg = ("#C94C4C", "#993030")
    err_hover = ("#D95C5C", "#A94040")

    active = {"on": True}
    start_t = time.monotonic()

    def apply_safe(**kw: Any) -> None:
        if not _button_alive(button):
            return
        try:
            button.configure(**kw)
        except Exception:
            pass

    def tick() -> None:
        if not active["on"]:
            return
        if not _button_alive(button):
            active["on"] = False
            return
        now = time.monotonic()
        apply_safe(text=f"{running_label} ({_format_elapsed(now - start_t)})")
        root.after(tick_ms, tick)

    def finish(rc: Optional[int], exc: Optional[BaseException], elapsed: float) -> None:
        active["on"] = False
        set_running(False)
        if not _button_alive(button):
            return

        et = _format_elapsed(elapsed)
        try:
            button.configure(state="normal")
        except Exception:
            pass

        if exc is not None:
            msg = str(exc).replace("\n", " ")
            if len(msg) > 48:
                msg = msg[:45] + "…"
            apply_safe(
                text=f"Ошибка: {msg}",
                fg_color=err_fg,
                hover_color=err_hover,
            )
        elif rc == 0:
            apply_safe(
                text=f"Готово ({et})",
                fg_color=ok_fg,
                hover_color=ok_hover,
            )
        else:
            apply_safe(
                text=f"Код выхода {rc} ({et})",
                fg_color=warn_fg,
                hover_color=warn_hover,
            )

        def reset() -> None:
            if not _button_alive(button):
                return
            try:
                kw: dict = {"text": base_text, "state": "normal"}
                if orig_fg is not None:
                    kw["fg_color"] = orig_fg
                if orig_hover is not None:
                    kw["hover_color"] = orig_hover
                button.configure(**kw)
            except Exception:
                pass

        root.after(reset_after_ms, reset)

    def thread_main() -> None:
        t0 = time.monotonic()
        rc: Optional[int] = None
        exc: Optional[BaseException] = None
        try:
            rc, exc = work_fn()
        except Exception as e:
            exc = e
        elapsed = time.monotonic() - t0
        root.after(0, lambda: finish(rc, exc, elapsed))

    try:
        apply_safe(
            state="disabled",
            fg_color=running_fg,
            hover_color=running_hover,
            text=f"{running_label} (0 с)",
        )
    except Exception:
        pass

    tick()
    threading.Thread(target=thread_main, daemon=True).start()


def _schedule_busy_flash(root: Any, button: Any) -> None:
    """Повторный клик во время работы: кратко «Уже выполняется…»; тикер перезапишет текст на следующем шаге."""
    if not _button_alive(button):
        return

    def flash() -> None:
        if not _button_alive(button):
            return
        try:
            button.configure(text="Уже выполняется…")
        except Exception:
            pass

    root.after(0, flash)


def attach_busy_aware_command(
    root: Any,
    button: Any,
    *,
    is_running: Callable[[], bool],
    start_job: Callable[[], None],
) -> None:
    def on_click() -> None:
        if is_running():
            _schedule_busy_flash(root, button)
            return
        start_job()

    try:
        button.configure(command=on_click)
    except Exception:
        pass
