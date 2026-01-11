passphrase en windows: willy.

Estoy trabajando con clave ssh-ed25519

# variosPython
varios scripts de python con herramientas

1.- Prueba de la librería click
2.- Version. Compara dos carpetas y si coinciden archivos, indica cual es el más actualizado
3.- Creación de cv en documento .docx
4.- Creación de cv en inglés .docx

**texto_comun.json   (vscode = alt+z) para ver el texo en un archivo json en varias lineas** no es permanente
Para hacerlo permanente --> 
Activación global (GUI):

Abre Settings (Ctrl+,), busca "Word Wrap" y cambia "Editor: Word Wrap" a "on".
Activación editando settings (global o workspace):

Para espacio de trabajo:
crear/editar .vscode/settings.json y añade:
para cualquier archivo
{
  "editor.wordWrap": "on"
}

solo para archivos json
{
  "[json]": {
    "editor.wordWrap": "on"
  }
}