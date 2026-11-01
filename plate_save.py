import os
import win32com.client

SAVE_DIR = r"C:\sw_project\parts"


def make_plate(name, length, width, height, n_holes, hole_d, step):
    """Строит пластину с рядом отверстий и сохраняет её. Размеры в метрах."""
    sw = win32com.client.Dispatch("SldWorks.Application")
    sw.Visible = True

    # Новая деталь
    template = sw.GetUserPreferenceStringValue(8)
    model = sw.NewDocument(template, 0, 0, 0)

    # Находим первую плоскость
    feat = model.FirstFeature
    plane = None
    while feat is not None:
        if feat.GetTypeName2 == "RefPlane":
            plane = feat
            break
        feat = feat.GetNextFeature

    # Эскиз: прямоугольник и ряд кругов
    plane.Select2(False, 0)
    model.SketchManager.InsertSketch(True)
    model.SketchManager.CreateCenterRectangle(0, 0, 0, length / 2, width / 2, 0)

    start = -step * (n_holes - 1) / 2
    for i in range(n_holes):
        x = start + i * step
        model.SketchManager.CreateCircleByRadius(x, 0, 0, hole_d / 2)

    # Выдавливание на толщину
    model.FeatureManager.FeatureExtrusion2(True, False, False, 0, 0, height, 0.01, False, False, False, False, 0.0174532925199433, 0.0174532925199433, False, False, False, False, True, True, True, 0, 0, False)

    model.ViewZoomtofit2

    # Сохранение в файл
    os.makedirs(SAVE_DIR, exist_ok=True)
    path = os.path.join(SAVE_DIR, name + ".SLDPRT")
    error = model.SaveAs3(path, 0, 1)
    if error == 0:
        print("Сохранено:", path)
    else:
        print("Не удалось сохранить, код ошибки:", error)


# Запуск: имя, длина, ширина, толщина, число отверстий, диаметр, шаг
make_plate("plate_1", 0.08, 0.05, 0.01, 3, 0.008, 0.02)
make_plate("plate_2", 0.1, 0.06, 0.008, 5, 0.006, 0.015)
print("Готово")