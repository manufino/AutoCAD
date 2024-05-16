
# AutoCAD Python Automation

Questa libreria Python fornisce una serie di classi e metodi per interagire con AutoCAD utilizzando l'API COM. La libreria permette di automatizzare molte operazioni comuni in AutoCAD, come la creazione e la gestione di layer, oggetti, blocchi, attributi e gruppi di oggetti.

## Funzionalità Principali

- **Gestione degli Strati (Layer)**: Creazione, modifica, impostazione della visibilità, blocco/sblocco, cambio del colore e gestione dello stile di linea dei layer.
- **Gestione degli Oggetti**: Creazione, selezione, spostamento, scalatura, rotazione, allineamento e distribuzione di oggetti.
- **Gestione dei Blocchi**: Inserimento, esportazione, creazione, modifica e rimozione di blocchi.
- **Gestione degli Attributi**: Aggiunta, modifica e cancellazione di attributi nei blocchi.
- **Input e Output dell'Utente**: Richiesta di input da parte dell'utente (punti, stringhe, numeri interi) e visualizzazione di messaggi.
- **Gestione di Gruppi di Oggetti**: Creazione, modifica, aggiunta/rimozione di oggetti e selezione di gruppi.

## Requisiti

- AutoCAD installato sul sistema.
- Python 3.x.
- pacchetto `pywin32` installato (installabile tramite pip).

## Installazione

1. Clona questo repository:
   ```sh
   git clone https://github.com/tuo-utente/autocad-python-automation.git
   ```

2. Installa le dipendenze:
   ```sh
   pip install pywin32
   ```

## Esempi di Utilizzo

Di seguito sono riportati alcuni esempi di utilizzo della libreria per automatizzare operazioni in AutoCAD.

## Creazione dell'oggetto AutoCAD

```python
# Creazione dell'oggetto AutoCAD
acad = AutoCAD()
```

## Creazione dei layer standard per il disegno meccanico

```python
# Creazione dei layer standard per il disegno meccanico
acad.create_standard_layers()
```

## Ripetizione del blocco "piatto" orizzontalmente

```python
# Ripetizione del blocco "piatto" orizzontalmente
total_length = 100  # Lunghezza totale X
block_length = 10  # Lunghezza del blocco "piatto"
insertion_point = APoint(0, 0, 0)  # Punto di inserimento iniziale

# Esegui la ripetizione del blocco
acad.repeat_block_horizontally("piatto", total_length, block_length, insertion_point)
```

## Imposta la visibilità di un layer

```python
# Imposta la visibilità di un layer
acad.set_layer_visibility("Linea di mezzeria", visible=False)
```

## Blocca un layer

```python
# Blocca un layer
acad.lock_layer("Quote", lock=True)
```

## Elimina un layer

```python
# Elimina un layer
acad.delete_layer("Simboli")
```

## Cambio del colore di un layer

```python
# Cambio del colore di un layer
acad.change_layer_color("Contorni", Color.YELLOW)
```

## Gestione dello stile di linea del layer

```python
# Gestione dello stile di linea del layer
acad.set_layer_linetype("Assi", "DASHED")
```

## Selezione di oggetti

```python
# Selezione di oggetti
selected_objects = acad.select_objects(object_type="AcDbLine", layer_name="Contorni")
print(f"Oggetti selezionati: {len(selected_objects)}")
```

## Sposta, scala e ruota oggetti

```python
# Sposta, scala e ruota oggetti
for obj in selected_objects:
    acad.move_object(obj, APoint(10, 10, 0))
    acad.scale_object(obj, APoint(0, 0, 0), 2)
    acad.rotate_object(obj, APoint(0, 0, 0), 45)
```

## Allineamento e distribuzione di oggetti

```python
# Allineamento e distribuzione di oggetti
acad.align_objects(selected_objects, alignment="left")
acad.distribute_objects(selected_objects, spacing=5)
```

## Inserimento di blocchi da file

```python
# Inserimento di blocchi da file
acad.insert_block_from_file("path_to_file.dwg", APoint(0, 0, 0))
```

## Esportazione di blocchi

```python
# Esportazione di blocchi
acad.export_block_to_file("piatto", "path_to_export.dwg")
```

## Modifica degli attributi

```python
# Modifica degli attributi
block_references = acad.get_block_coordinates("piatto")
if block_references:
    block_ref = block_references[0]  # Prendi il primo blocco trovato
    acad.modify_block_attribute(block_ref, "Tag", "NewValue")
```

## Cancellazione degli attributi

```python
# Cancellazione degli attributi
acad.delete_block_attribute(block_ref, "Tag")
```

## Richieste di input da utente

```python
# Richieste di input da utente
point = acad.get_user_input_point("Seleziona un punto")
text = acad.get_user_input_string("Inserisci un testo")
integer = acad.get_user_input_integer("Inserisci un numero intero")
```

## Messaggi di output all'utente

```python
# Messaggi di output all'utente
acad.show_message("Operazione completata")
```

## Creazione di gruppi di oggetti

```python
# Creazione di gruppi di oggetti
group = acad.create_group("MyGroup", selected_objects)
```

## Aggiungi oggetti a un gruppo

```python
# Aggiungi oggetti a un gruppo
acad.add_to_group("MyGroup", selected_objects)
```

## Rimuovi oggetti da un gruppo

```python
# Rimuovi oggetti da un gruppo
acad.remove_from_group("MyGroup", selected_objects)
```

## Seleziona gruppi di oggetti

```python
# Seleziona gruppi di oggetti
group_items = acad.select_group("MyGroup")
print(f"Oggetti nel gruppo 'MyGroup': {len(group_items)}")
```

## Stampa i layer creati per conferma

```python
# Stampa i layer creati per conferma
for layer in acad.doc.Layers:
    print(f"Layer: {layer.Name}, Colore: {layer.color}")
```
