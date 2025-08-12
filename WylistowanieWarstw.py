import arcpy
import csv

def main():
    # Parametry wejściowe
    output_csv = arcpy.GetParameterAsText(0)  # Ścieżka do pliku CSV
    selected_type = arcpy.GetParameterAsText(1)  # Typ warstw do filtrowania

    # Pobierz aktywną mapę
    aprx = arcpy.mp.ArcGISProject("CURRENT")
    active_map = aprx.activeMap

    if not active_map:
        arcpy.AddError("Brak aktywnej mapy w projekcie.")
        return

    all_layers = []

    # Przejście po warstwach
    for lyr in active_map.listLayers():
        try:
            desc = arcpy.Describe(lyr)
            data_type = desc.dataType

            # Filtrowanie po typie warstwy
            if selected_type != "Wszystkie" and data_type != selected_type:
                continue

            full_path = lyr.dataSource if hasattr(lyr, "dataSource") else ""
            geom_type = getattr(desc, "shapeType", "")
            count = ""
            try:
                count = int(arcpy.GetCount_management(lyr)[0])
            except:
                pass
            sr = desc.spatialReference.name if hasattr(desc, "spatialReference") else ""

            all_layers.append([lyr.name, full_path, data_type, geom_type, count, sr])
            
            arcpy.AddMessage("\n\n-------------------------------------------------------------------------------------------------------------------------------")
            arcpy.AddMessage(f"\n\nNazwa warstwy: {lyr.name}\n\nŚcieżka: {full_path}\n\nTyp danych: {data_type}\n\nTyp geometrii: {geom_type}\n\nLiczba rekordów: {count}\n\nUkład współrzędnych: {sr}")
           
        except Exception as e:
            arcpy.AddWarning(f"Nie udało się z warstwą '{lyr.name}': {e}")

    # Przejście po tabelach
    for table in active_map.listTables():
        try:
            desc = arcpy.Describe(table)
            data_type = desc.dataType

            if selected_type != "Wszystkie" and data_type != selected_type:
                continue

            full_path = table.dataSource if hasattr(table, "dataSource") else ""
            count = ""
            try:
                count = int(arcpy.GetCount_management(table)[0])
            except:
                pass
            sr = desc.spatialReference.name if hasattr(desc, "spatialReference") else ""

            all_layers.append([table.name, full_path, data_type, "", count, sr])
           
            arcpy.AddMessage("\n\n-------------------------------------------------------------------------------------------------------------------------------")
            arcpy.AddMessage(f"\n\nNazwa tabeli: {table.name}\n\nŚcieżka: {full_path}\n\nTyp danych: {data_type}\n\nLiczba rekordów: {count}\n\nUkład współrzędnych: {sr}")
            
        except Exception as e:
            arcpy.AddWarning(f"Nie udało się z tabelą '{table.name}': {e}")

    # Zapis do pliku CSV
    try:
        with open(output_csv, mode="w", newline="", encoding="utf-8-sig") as file: #kodowanie utf-8-sig (UTF-8 BOM bardzo ważne bo csv excel nie czyta polskich znaków inaczej :D)
            writer = csv.writer(file, delimiter=";")
            writer.writerow(["Layer name", "Path", "Data type", "Geometry type", "Record count", "Spatial Reference"])
            writer.writerows(all_layers)
        arcpy.AddMessage("-" * 160)
        arcpy.AddMessage(f"Wygenerowano listę warstw i tabel: {output_csv}")
        arcpy.AddMessage("-" * 160)
    except Exception as e:
        arcpy.AddMessage("-" * 160)
        arcpy.AddError(f"Nie udało się zapisać pliku CSV: {e}")
        arcpy.AddMessage("-" * 160)
if __name__ == "__main__":
    main()


