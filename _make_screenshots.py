"""
Утиліта для створення скріншотів програми для README.
Запускати один раз перед оновленням документації:
    python _make_screenshots.py
"""
import os, sys, time
from PIL import ImageGrab
import tkinter as tk

import bron_analysis_all as m

OUT = os.path.join(os.path.dirname(__file__), "screenshots")
os.makedirs(OUT, exist_ok=True)


def grab_window(root, filename):
    """Скріншот вікна root з полями."""
    root.update_idletasks()
    root.update()
    time.sleep(0.5)
    x = root.winfo_rootx()
    y = root.winfo_rooty() - 30  # title bar
    w = root.winfo_width()
    h = root.winfo_height() + 30
    img = ImageGrab.grab(bbox=(x, y, x + w, y + h))
    path = os.path.join(OUT, filename)
    img.save(path)
    print(f"  -> {filename}  ({img.size[0]}x{img.size[1]})")


def main():
    app = m.App()
    app.geometry("1280x820+50+20")
    app.update()
    time.sleep(0.5)

    # Завантажуємо приклади
    here = os.path.dirname(os.path.abspath(__file__))
    app._load_file(os.path.join(here, "sample_zp.xlsx"))
    app._analyse_sheet("Регістер")
    app._load_dia_file_silent(os.path.join(here, "sample_dia.xlsx"))
    app.update(); time.sleep(0.8)

    # 1. Головна вкладка
    app.nb.select(app.tab_main)
    app.update(); time.sleep(0.6)
    grab_window(app, "01-main-table.png")

    # 2. Підсумки
    app.nb.select(app.tab_summary)
    app.update(); time.sleep(0.6)
    grab_window(app, "02-summary.png")

    # 3. Аналітика
    app.nb.select(app.tab_avg)
    app.update(); time.sleep(0.6)
    grab_window(app, "03-analytics.png")

    # 4. Легенда кольорів
    app.nb.select(app.tab_main)
    app.update(); time.sleep(0.3)
    app._show_legend()
    app.update(); time.sleep(0.6)
    grab_window(app, "04-legend.png")
    for w in app.winfo_children():
        if isinstance(w, tk.Toplevel):
            w.destroy()

    # 5. Деталі працівника
    app.nb.select(app.tab_main)
    app.update(); time.sleep(0.3)
    # Знайдемо першого заброньованого з даними
    for item in app.tree.get_children():
        vals = app.tree.item(item, "values")
        if vals and len(vals) >= 4 and str(vals[3]) == "Так":
            app.tree.selection_set(item)
            app._show_employee_details(vals[1])
            break
    app.update(); time.sleep(0.6)
    grab_window(app, "05-employee-details.png")
    for w in app.winfo_children():
        if isinstance(w, tk.Toplevel):
            w.destroy()

    app.destroy()
    print(f"\nГотово! Скріншоти збережено в: {OUT}")


if __name__ == "__main__":
    main()
