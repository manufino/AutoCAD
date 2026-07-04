"""
Example script for AutoCAD.py: draws a small house scene (walls, roof,
door, windows, a sun and a tree), adds a dimension for the house width,
and labels the drawing.

Run it with AutoCAD already open on a drawing - AutoCAD() attaches to
the running instance by default.
"""

import math

from AutoCAD import AutoCAD, APoint, Color, Layer, Text, Dimension, DimensionType


def draw_house(acad, base_point):
    x0, y0 = base_point.x, base_point.y

    acad.create_layer(Layer("Struttura", color=Color.WHITE))
    acad.set_active_layer("Struttura")
    acad.add_rectangle(APoint(x0, y0), APoint(x0 + 200, y0 + 150))
    acad.add_rectangle(APoint(x0 + 80, y0), APoint(x0 + 120, y0 + 70))  # door
    acad.add_rectangle(APoint(x0 + 20, y0 + 90), APoint(x0 + 60, y0 + 130))  # window
    acad.add_rectangle(APoint(x0 + 140, y0 + 90), APoint(x0 + 180, y0 + 130))  # window

    acad.create_layer(Layer("Tetto", color=Color.RED))
    acad.set_active_layer("Tetto")
    acad.add_polyline([
        APoint(x0 - 20, y0 + 150),
        APoint(x0 + 100, y0 + 230),
        APoint(x0 + 220, y0 + 150),
        APoint(x0 - 20, y0 + 150),
    ])


def draw_sun(acad, center, radius):
    acad.create_layer(Layer("Sole", color=Color.YELLOW))
    acad.set_active_layer("Sole")
    acad.add_circle(center, radius)

    for i in range(8):
        angle = 2 * math.pi * i / 8
        start = APoint(center.x + radius * 1.3 * math.cos(angle), center.y + radius * 1.3 * math.sin(angle))
        end = APoint(center.x + radius * 1.7 * math.cos(angle), center.y + radius * 1.7 * math.sin(angle))
        acad.add_line(start, end)


def draw_tree(acad, base_point):
    x0, y0 = base_point.x, base_point.y

    acad.create_layer(Layer("Albero", color=Color.GREEN))
    acad.set_active_layer("Albero")
    acad.add_rectangle(APoint(x0 - 10, y0), APoint(x0 + 10, y0 + 60))  # trunk

    # Foliage: three overlapping circles for a fluffy look
    acad.add_circle(APoint(x0 - 20, y0 + 85), 30)
    acad.add_circle(APoint(x0 + 20, y0 + 85), 30)
    acad.add_circle(APoint(x0, y0 + 110), 30)


def draw_ground_and_title(acad, y0=0):
    acad.create_layer(Layer("Terreno", color=Color.GRAY))
    acad.set_active_layer("Terreno")
    acad.add_line(APoint(-250, y0), APoint(400, y0))

    acad.create_layer(Layer("Testo", color=Color.WHITE))
    acad.set_active_layer("Testo")
    acad.add_text(Text("Disegnato con AutoCAD.py", APoint(0, y0 - 40), 15))


def draw_dimension(acad, base_point):
    acad.create_layer(Layer("Quote", color=Color.CYAN))
    acad.set_active_layer("Quote")
    acad.add_dimension(Dimension(
        base_point, APoint(base_point.x + 200, base_point.y), APoint(base_point.x, base_point.y - 20),
        dimension_type=DimensionType.ALIGNED,
    ))


def main():
    acad = AutoCAD()  # attaches to the already-open drawing

    house_base = APoint(0, 0)
    draw_ground_and_title(acad, y0=house_base.y)
    draw_house(acad, house_base)
    draw_dimension(acad, house_base)
    draw_sun(acad, APoint(300, 250), 30)
    draw_tree(acad, APoint(-140, 0))

    try:
        acad.acad.ZoomExtents()
    except Exception:
        pass  # zoom is a nice-to-have, not critical

    print("Disegno completato in AutoCAD.")


if __name__ == "__main__":
    main()
