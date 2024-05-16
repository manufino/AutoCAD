import win32com.client
import pythoncom
from enum import Enum

# Enum per rappresentare i colori comuni in AutoCAD
class Color(Enum):
    RED = 1
    YELLOW = 2
    GREEN = 3
    CYAN = 4
    BLUE = 5
    MAGENTA = 6
    WHITE = 7
    GRAY = 8
    ORANGE = 30
    PURPLE = 40
    BROWN = 41

    @staticmethod
    def from_name(name):
        try:
            return Color[name.upper()].value
        except KeyError:
            raise ValueError(f"Color '{name}' is not a valid color name")

# Classe per rappresentare un punto 3D
class APoint:
    def __init__(self, x=0, y=0, z=0):
        self.x = x
        self.y = y
        self.z = z

    def to_variant(self):
        return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [self.x, self.y, self.z])

    def to_tuple(self):
        return (self.x, self.y)

    def __repr__(self):
        return f"APoint({self.x}, {self.y}, {self.z})"

# Classe per rappresentare un layer
class Layer:
    def __init__(self, name, color=Color.WHITE, visible=True):
        self.name = name
        self.color = color
        self.visible = visible

    def __repr__(self):
        return f"Layer(name='{self.name}', color={self.color}, visible={self.visible})"

# Classe per rappresentare un riferimento di blocco
class BlockReference:
    def __init__(self, name, insertion_point, scale=1.0, rotation=0.0):
        self.name = name
        self.insertion_point = insertion_point
        self.scale = scale
        self.rotation = rotation

    def __repr__(self):
        return f"BlockReference(name='{self.name}', insertion_point={self.insertion_point}, scale={self.scale}, rotation={self.rotation})"

# Classe per rappresentare un testo
class Text:
    def __init__(self, content, insertion_point, height, alignment='left'):
        self.content = content
        self.insertion_point = insertion_point
        self.height = height
        self.alignment = alignment

    def __repr__(self):
        return f"Text(content='{self.content}', insertion_point={self.insertion_point}, height={self.height}, alignment='{self.alignment}')"

# Classe per rappresentare una quota
class Dimension:
    def __init__(self, start_point, end_point, text_point, dimension_type='aligned'):
        self.start_point = start_point
        self.end_point = end_point
        self.text_point = text_point
        self.dimension_type = dimension_type

    def __repr__(self):
        return f"Dimension(start_point={self.start_point}, end_point={self.end_point}, text_point={self.text_point}, dimension_type='{self.dimension_type}')"

# Classe personalizzata per gestire le eccezioni di AutoCAD
class AutoCADError(Exception):
    def __init__(self, message):
        super().__init__(message)
        # Puoi aggiungere qui ulteriori logiche di gestione dell'errore, come il logging su file
        print(f"AutoCADError: {message}")

# Classe principale per interagire con AutoCAD
class AutoCAD:
    def __init__(self):
        try:
            self.acad = win32com.client.Dispatch("AutoCAD.Application")
            self.acad.Visible = True
            self.doc = self.acad.ActiveDocument
            self.modelspace = self.doc.ModelSpace
        except Exception as e:
            raise AutoCADError(f"Errore durante l'inizializzazione di AutoCAD: {e}")

    # Itera sugli oggetti nello spazio modello, opzionalmente filtrando per tipo di oggetto
    def iter_objects(self, object_type=None):
        for obj in self.modelspace:
            if object_type is None or obj.EntityName == object_type:
                yield obj

    # Aggiunge un cerchio nello spazio modello
    def add_circle(self, center, radius):
        try:
            circle = self.modelspace.AddCircle(center.to_variant(), radius)
            return circle
        except Exception as e:
            raise AutoCADError(f"Errore durante l'aggiunta di un cerchio: {e}")

    # Aggiunge una linea nello spazio modello
    def add_line(self, start_point, end_point):
        try:
            line = self.modelspace.AddLine(start_point.to_variant(), end_point.to_variant())
            return line
        except Exception as e:
            raise AutoCADError(f"Errore durante l'aggiunta di una linea: {e}")

    # Aggiunge un rettangolo nello spazio modello
    def add_rectangle(self, lower_left, upper_right):
        try:
            x1, y1 = lower_left.to_tuple()
            x2, y2 = upper_right.to_tuple()
            points = [
                x1, y1,
                x2, y1,
                x2, y2,
                x1, y2,
                x1, y1
            ]
            points_variant = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, points)
            polyline = self.modelspace.AddLightweightPolyline(points_variant)
            return polyline
        except Exception as e:
            raise AutoCADError(f"Errore durante l'aggiunta di un rettangolo: {e}")

    # Aggiunge un'ellisse nello spazio modello
    def add_ellipse(self, center, major_axis, ratio):
        try:
            ellipse = self.modelspace.AddEllipse(center.to_variant(), major_axis.to_variant(), ratio)
            return ellipse
        except Exception as e:
            raise AutoCADError(f"Errore durante l'aggiunta di un'ellisse: {e}")

    # Aggiunge un testo nello spazio modello
    def add_text(self, text):
        try:
            text_obj = self.modelspace.AddText(text.content, text.insertion_point.to_variant(), text.height)
            return text_obj
        except Exception as e:
            raise AutoCADError(f"Errore durante l'aggiunta di un testo: {e}")

    # Aggiunge una quota nello spazio modello
    def add_dimension(self, dimension):
        try:
            dimension_obj = None
            if dimension.dimension_type == 'aligned':
                dimension_obj = self.modelspace.AddDimAligned(dimension.start_point.to_variant(), dimension.end_point.to_variant(), dimension.text_point.to_variant())
            return dimension_obj
        except Exception as e:
            raise AutoCADError(f"Errore durante l'aggiunta di una quota: {e}")

    # Ottiene i blocchi definiti dall'utente nel documento
    def get_user_defined_blocks(self):
        try:
            blocks = self.doc.Blocks
            user_defined_blocks = [block.Name for block in blocks 
                                   if not block.IsLayout and not block.Name.startswith('*') and block.Name != 'GENAXEH']
            return user_defined_blocks
        except Exception as e:
            raise AutoCADError(f"Errore durante il recupero dei blocchi definiti dall'utente: {e}")

    # Crea un nuovo layer
    def create_layer(self, layer):
        try:
            layers = self.doc.Layers
            new_layer = layers.Add(layer.name)
            new_layer.Color = layer.color.value
            return new_layer
        except Exception as e:
            raise AutoCADError(f"Errore durante la creazione del layer '{layer.name}': {e}")

    # Imposta il layer attivo
    def set_active_layer(self, layer_name):
        try:
            self.doc.ActiveLayer = self.doc.Layers.Item(layer_name)
        except Exception as e:
            raise AutoCADError(f"Errore durante l'impostazione del layer attivo '{layer_name}': {e}")

    # Inserisce un blocco nello spazio modello
    def insert_block(self, block):
        try:
            block_ref = self.modelspace.InsertBlock(block.insertion_point.to_variant(), block.name, block.scale, block.scale, block.scale, block.rotation)
            return block_ref
        except Exception as e:
            raise AutoCADError(f"Errore durante l'inserimento del blocco '{block.name}': {e}")

    # Salva il documento con un nuovo nome
    def save_as(self, file_path):
        try:
            self.doc.SaveAs(file_path)
        except Exception as e:
            raise AutoCADError(f"Errore durante il salvataggio del documento come '{file_path}': {e}")

    # Apre un file esistente
    def open_file(self, file_path):
        try:
            self.acad.Documents.Open(file_path)
        except Exception as e:
            raise AutoCADError(f"Errore durante l'apertura del file '{file_path}': {e}")

    # Ottiene le coordinate di inserimento di un blocco specifico
    def get_block_coordinates(self, block_name):
        try:
            block_references = []
            for entity in self.iter_objects("AcDbBlockReference"):
                if entity.Name == block_name:
                    insertion_point = entity.InsertionPoint
                    block_references.append(APoint(insertion_point[0], insertion_point[1], insertion_point[2]))
            return block_references
        except Exception as e:
            raise AutoCADError(f"Errore durante il recupero delle coordinate del blocco '{block_name}': {e}")

    # Elimina un oggetto
    def delete_object(self, obj):
        try:
            obj.Delete()
        except Exception as e:
            raise AutoCADError(f"Errore durante l'eliminazione dell'oggetto: {e}")

    # Clona un oggetto
    def clone_object(self, obj, new_insertion_point):
        try:
            cloned_obj = obj.Copy(new_insertion_point.to_variant())
            return cloned_obj
        except Exception as e:
            raise AutoCADError(f"Errore durante la clonazione dell'oggetto: {e}")

    # Modifica una proprietà di un oggetto
    def modify_object_property(self, obj, property_name, new_value):
        try:
            setattr(obj, property_name, new_value)
        except Exception as e:
            raise AutoCADError(f"Errore durante la modifica della proprietà '{property_name}' dell'oggetto: {e}")

    # Crea una serie di layer standard per il disegno meccanico
    def create_standard_layers(self):
        standard_layers = [
            {"name": "Linea di mezzeria", "color": Color.RED},
            {"name": "Quote", "color": Color.GREEN},
            {"name": "Contorni", "color": Color.BLUE},
            {"name": "Assi", "color": Color.CYAN},
            {"name": "Testi", "color": Color.MAGENTA},
            {"name": "Simboli", "color": Color.YELLOW},
            {"name": "Tratteggi", "color": Color.ORANGE},
            {"name": "Costruzione", "color": Color.PURPLE}
        ]

        for layer in standard_layers:
            self.create_layer(Layer(layer["name"], layer["color"]))

    # Ripete orizzontalmente un blocco fino a raggiungere la lunghezza specificata
    def repeat_block_horizontally(self, block_name, total_length, block_length, insertion_point):
        try:
            x, y, z = insertion_point.x, insertion_point.y, insertion_point.z
            num_blocks = total_length // block_length

            for i in range(int(num_blocks)):
                new_insertion_point = APoint(x + i * block_length, y, z)
                self.insert_block(BlockReference(block_name, new_insertion_point))
        except Exception as e:
            raise AutoCADError(f"Errore durante la ripetizione del blocco '{block_name}' orizzontalmente: {e}")

    # Imposta la visibilità di un layer
    def set_layer_visibility(self, layer_name, visible=True):
        try:
            layer = self.doc.Layers.Item(layer_name)
            layer.LayerOn = visible
        except Exception as e:
            raise AutoCADError(f"Errore durante l'impostazione della visibilità del layer '{layer_name}': {e}")

    # Blocca o sblocca un layer
    def lock_layer(self, layer_name, lock=True):
        try:
            layer = self.doc.Layers.Item(layer_name)
            layer.Lock = lock
        except Exception as e:
            raise AutoCADError(f"Errore durante il blocco/sblocco del layer '{layer_name}': {e}")

    # Elimina un layer
    def delete_layer(self, layer_name):
        try:
            layer = self.doc.Layers.Item(layer_name)
            layer.Delete()
        except Exception as e:
            raise AutoCADError(f"Errore durante l'eliminazione del layer '{layer_name}': {e}")

    # Cambio del colore di un layer
    def change_layer_color(self, layer_name, color):
        try:
            layer = self.doc.Layers.Item(layer_name)
            layer.color = color.value
        except Exception as e:
            raise AutoCADError(f"Errore durante il cambio del colore del layer '{layer_name}': {e}")

    # Gestione dello stile di linea del layer
    def set_layer_linetype(self, layer_name, linetype_name):
        try:
            layer = self.doc.Layers.Item(layer_name)
            linetypes = self.doc.Linetypes
            if linetype_name not in linetypes:
                self.doc.Linetypes.Load(linetype_name, linetype_name)
            layer.Linetype = linetype_name
        except Exception as e:
            raise AutoCADError(f"Errore durante l'impostazione dello stile di linea del layer '{layer_name}': {e}")

    # Sposta, scala e ruota oggetti
    def move_object(self, obj, new_insertion_point):
        try:
            obj.Move(obj.InsertionPoint, new_insertion_point.to_variant())
        except Exception as e:
            raise AutoCADError(f"Errore durante lo spostamento dell'oggetto: {e}")

    def scale_object(self, obj, base_point, scale_factor):
        try:
            obj.ScaleEntity(base_point.to_variant(), scale_factor)
        except Exception as e:
            raise AutoCADError(f"Errore durante la scalatura dell'oggetto: {e}")

    def rotate_object(self, obj, base_point, rotation_angle):
        try:
            obj.Rotate(base_point.to_variant(), rotation_angle)
        except Exception as e:
            raise AutoCADError(f"Errore durante la rotazione dell'oggetto: {e}")

    # Allineamento e distribuzione di oggetti
    def align_objects(self, objects, alignment="left"):
        try:
            if not objects:
                return
            if alignment == "left":
                min_x = min(obj.InsertionPoint[0] for obj in objects)
                for obj in objects:
                    self.move_object(obj, APoint(min_x, obj.InsertionPoint[1], obj.InsertionPoint[2]))
            elif alignment == "right":
                max_x = max(obj.InsertionPoint[0] for obj in objects)
                for obj in objects:
                    self.move_object(obj, APoint(max_x, obj.InsertionPoint[1], obj.InsertionPoint[2]))
        except Exception as e:
            raise AutoCADError(f"Errore durante l'allineamento degli oggetti: {e}")

    def distribute_objects(self, objects, spacing):
        try:
            if not objects:
                return
            objects.sort(key=lambda obj: obj.InsertionPoint[0])
            for i in range(1, len(objects)):
                new_x = objects[i-1].InsertionPoint[0] + spacing
                self.move_object(objects[i], APoint(new_x, objects[i].InsertionPoint[1], objects[i].InsertionPoint[2]))
        except Exception as e:
            raise AutoCADError(f"Errore durante la distribuzione degli oggetti: {e}")

    # Inserimento di blocchi da file
    def insert_block_from_file(self, file_path, insertion_point, scale=1.0, rotation=0.0):
        try:
            block_name = self.doc.Blocks.Import(file_path, file_path)
            block_ref = self.modelspace.InsertBlock(insertion_point.to_variant(), block_name, scale, scale, scale, rotation)
            return block_ref
        except Exception as e:
            raise AutoCADError(f"Errore durante l'inserimento del blocco da file '{file_path}': {e}")

    # Esportazione di blocchi
    def export_block_to_file(self, block_name, file_path):
        try:
            block = self.doc.Blocks.Item(block_name)
            block.Export(file_path)
        except Exception as e:
            raise AutoCADError(f"Errore durante l'esportazione del blocco '{block_name}' in '{file_path}': {e}")

    # Modifica degli attributi
    def modify_block_attribute(self, block_ref, tag, new_value):
        try:
            for attribute in block_ref.GetAttributes():
                if attribute.TagString == tag:
                    attribute.TextString = new_value
        except Exception as e:
            raise AutoCADError(f"Errore durante la modifica dell'attributo '{tag}' del blocco: {e}")

    # Cancellazione degli attributi
    def delete_block_attribute(self, block_ref, tag):
        try:
            for attribute in block_ref.GetAttributes():
                if attribute.TagString == tag:
                    attribute.Delete()
        except Exception as e:
            raise AutoCADError(f"Errore durante la cancellazione dell'attributo '{tag}' del blocco: {e}")

    # Richieste di input da utente
    def get_user_input_point(self, prompt="Seleziona un punto"):
        try:
            point = self.doc.Utility.GetPoint(None, prompt)
            return APoint(point[0], point[1], point[2])
        except Exception as e:
            raise AutoCADError(f"Errore durante la richiesta del punto all'utente: {e}")

    def get_user_input_string(self, prompt="Inserisci un testo"):
        try:
            return self.doc.Utility.GetString(False, prompt)
        except Exception as e:
            raise AutoCADError(f"Errore durante la richiesta della stringa all'utente: {e}")

    def get_user_input_integer(self, prompt="Inserisci un numero intero"):
        try:
            return self.doc.Utility.GetInteger(prompt)
        except Exception as e:
            raise AutoCADError(f"Errore durante la richiesta del numero intero all'utente: {e}")

    # Messaggi di output all'utente
    def show_message(self, message):
        try:
            self.doc.Utility.Prompt(message + "\n")
        except Exception as e:
            raise AutoCADError(f"Errore durante la visualizzazione del messaggio: {e}")

    # Creazione di gruppi di oggetti
    def create_group(self, group_name, objects):
        try:
            group = self.doc.Groups.Add(group_name)
            for obj in objects:
                group.AppendItems([obj])
            return group
        except Exception as e:
            raise AutoCADError(f"Errore durante la creazione del gruppo '{group_name}': {e}")

    # Aggiungi oggetti a un gruppo
    def add_to_group(self, group_name, objects):
        try:
            group = self.doc.Groups.Item(group_name)
            for obj in objects:
                group.AppendItems([obj])
        except Exception as e:
            raise AutoCADError(f"Errore durante l'aggiunta di oggetti al gruppo '{group_name}': {e}")

    # Rimuovi oggetti da un gruppo
    def remove_from_group(self, group_name, objects):
        try:
            group = self.doc.Groups.Item(group_name)
            for obj in objects:
                group.RemoveItems([obj])
        except Exception as e:
            raise AutoCADError(f"Errore durante la rimozione di oggetti dal gruppo '{group_name}': {e}")

    # Seleziona gruppi di oggetti
    def select_group(self, group_name):
        try:
            group = self.doc.Groups.Item(group_name)
            return [item for item in group.GetItems()]
        except Exception as e:
            raise AutoCADError(f"Errore durante la selezione del gruppo '{group_name}': {e}")

