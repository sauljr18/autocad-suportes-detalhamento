import threading
import time

import pythoncom
import win32com.client

import flet as ft


def APoint(x, y, z=0):
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (x, y, z))

def aDouble(xyz):
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (xyz))

def aDispatch(vObject):
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, vObject)

def inicializar_acad():
    global acad, acadDoc, acadModel

    acad = None
    try:
        # Tentativa para conectar ao AutoCAD já aberto
        acad = win32com.client.GetActiveObject("AutoCAD.Application")
    except Exception as e:
        # Se não conseguir, tenta inicializar uma nova instância
        try:
            acad = win32com.client.dynamic.Dispatch("AutoCAD.Application", pythoncom.CoInitialize())
        except Exception as e:
            print(f"Erro ao conectar ao AutoCAD: {e}")
            return None, None, None

    if acad:
        while acad.Documents.Count == 0:
            print("Aguardando um documento ativo no AutoCAD...")
            time.sleep(2)

        acadDoc = acad.ActiveDocument
        acadModel = acadDoc.ModelSpace

        print(f"Documento ativo: {acadDoc.Name}")
        return acad, acadDoc, acadModel

    print("Falha ao inicializar o AutoCAD.")
    return None, None, None

acad, acadDoc, acadModel = inicializar_acad()

def contar_valores(lista):
    contagem = {}
    for valor in lista:
        contagem[valor] = contagem.get(valor, 0) + 1
    return contagem

def modificar_atributos_bloco(acadDoc, acadModel):
    print('Quantidade de Blocos:', acadDoc.Blocks.Count)

    for entity in acadModel:
        if entity.EntityName == 'AcDbBlockReference':
            if entity.HasAttributes:
                print(f'Nome: {entity.Name}, Layer: {entity.Layer}, Object ID: {entity.ObjectID}')
                for attrib in entity.GetAttributes():
                    print(f"Atributo {attrib.TagString}: {attrib.TextString}")
                    if attrib.TagString == 'Title':
                        attrib.TextString = 'Modified Title'
                    attrib.Update()
            if entity.IsDynamicBlock:
                for dyn_prop in entity.GetDynamicBlockProperties():
                    print(f'Dynamic Property: {dyn_prop.PropertyName} = {dyn_prop.Value}')
                    if dyn_prop.PropertyName == 'MEDIDA H':
                        dyn_prop.Value = 237
                print('Bloco dinâmico modificado.')

    acad.ZoomExtents()

def zoom_center(acad, x, y, z):
    p1 = APoint(x-200, y+200, z)
    p2 = APoint(x+200, y-200, z)
    acad.ZoomWindow(p1, p2)

def main(page: ft.Page):
    page.window.full_screen = False
    page.title = "Suportes de Tubulação"

    # Visible debug banner to confirm page rendering
    page.add(ft.Container(content=ft.Text("DEBUG: página iniciada", size=16, weight=ft.FontWeight.BOLD), bgcolor=ft.Colors.PINK_ACCENT_100, padding=10))

    searchtxt = ft.TextField(label="Buscar", value="", keyboard_type="text")
    search_option = ft.RadioGroup(
        content=ft.Row([
            ft.Radio(value="nomeSuporte", label="Nome Suporte"),
            ft.Radio(value="tipoSuporte", label="Tipo Suporte")
        ]),
        value="nomeSuporte",
        on_change=""
    )

    selected_row = None

    def highlight_row(row):
        nonlocal selected_row
        if selected_row is not None:
            selected_row.color = None
        row.color = ft.Colors.YELLOW
        selected_row = row

    def zoom_to_block(e):
        bloco = e.control.data
        pos_x, pos_y, pos_z = float(bloco['Posicao X']), float(bloco['Posicao Y']), float(bloco['Posicao Z'])
        zoom_center(acad, pos_x, pos_y, pos_z)
        highlight_row(e.control.parent.parent.parent)
        page.update()

    def mostra_controle(e):
        row_controls = e.control.data["row_controls"]
        text_field, button_ok = row_controls["text_field"], row_controls["button_ok"]
        text_field.visible = True
        button_ok.visible = True
        highlight_row(e.control.parent.parent.parent)
        page.update()

    def listarPropriedades(e):
        handle = e.control.data['Handle']
        bloco_propriedades = {
            prop.PropertyName: prop.Value for entity in acadModel
            if entity.EntityName == 'AcDbBlockReference' and entity.Handle == handle
            for prop in entity.GetDynamicBlockProperties() if prop.PropertyName != "Origin"
        }
        popula_dt_blocos_valores(dict(sorted(bloco_propriedades.items())), handle)
        highlight_row(e.control.parent.parent.parent)
        container_dt_blocos_valores.visible = True
        page.update()

    def popula_dt_blocos_valores(bloco_propriedades, handle):
        dt_blocos_valores.rows.clear()
        for atributo, valor in bloco_propriedades.items():
            text_field = ft.TextField(label="Novo Valor", value="", visible=False, width=120)
            button_ok = ft.IconButton(
                icon=ft.Icons.CHECK_CIRCLE_ROUNDED, visible=False,
                data={'Atributo': atributo, 'Handle': handle, 'NovoValor': ""},
                on_click=lambda e, text_field=text_field: atualiza_valor_text_field(e, text_field)
            )
            dt_blocos_valores.rows.append(
                ft.DataRow(cells=[
                    ft.DataCell(ft.Text(str(atributo))),
                    ft.DataCell(ft.Text(str(valor))),
                    ft.DataCell(ft.Text(str(handle))),
                    ft.DataCell(ft.Container(
                        ft.Row([
                            ft.IconButton(icon=ft.Icons.EDIT, icon_color="green",
                                          data={'row_controls': {'text_field': text_field, 'button_ok': button_ok}},
                                          on_click=mostra_controle
                            ),
                            text_field, button_ok
                        ]), width=260
                    )),
                ])
            )
        page.update()

    def atualiza_valor_text_field(e, text_field):
        e.control.data["NovoValor"] = text_field.value
        atualiza_valor_propriedade(e)

    def atualiza_valor_propriedade(e):
        dados = e.control.data
        handle, nome_propriedade, novoValor = dados['Handle'], dados['Atributo'], float(dados['NovoValor'])
        for entity in acadModel:
            if entity.EntityName == 'AcDbBlockReference' and entity.Handle == handle:
                for prop in entity.GetDynamicBlockProperties():
                    if prop.PropertyName == nome_propriedade:
                        try:
                            if hasattr(prop, 'ValueMinimum') and hasattr(prop, 'ValueMaximum'):
                                if prop.ValueMinimum <= novoValor <= prop.ValueMaximum:
                                    prop.Value = float(novoValor)
                                    print(f"Novo valor de '{nome_propriedade}': {prop.Value}")
                                else:
                                    print("O valor desejado está fora dos limites permitidos.")
                            else:
                                prop.Value = float(novoValor)
                                print(f"Novo valor de '{nome_propriedade}': {prop.Value}")
                        except Exception as erro:
                            print(f"Erro ao tentar alterar o valor: {erro}")
        mensagemPop("Valor atualizado com sucesso!", 20, ft.Colors.GREEN)
        threading.Thread(target=hide_container_after_delay).start()

    def listarSuportes(e=None):
        acad, acadDoc, acadModel = inicializar_acad()
        if acadDoc is not None and acadModel is not None:
            print("AutoCAD inicializado com sucesso!")
            print('Quantidade de Blocos:', acadDoc.Blocks.Count)
        else:
            print("Não foi possível inicializar o AutoCAD ou obter um documento ativo.")

        mydt.rows.clear()
        valores_blocos = [
            {
                'Tag': attrib.TextString, 'Tipo': entity.Name,
                'Posicao X': attrib.InsertionPoint[0], 'Posicao Y': attrib.InsertionPoint[1],
                'Posicao Z': attrib.InsertionPoint[2], 'Handle': entity.Handle
            }
            for entity in acadModel if entity.EntityName == 'AcDbBlockReference' and entity.HasAttributes
            for attrib in entity.GetAttributes() if attrib.TagString == 'POSICAO'
        ]
        valores_blocos_ordenados = sorted(valores_blocos, key=lambda bloco: bloco['Tag'])
        for bloco in valores_blocos_ordenados:
            tag, tipo, pos_x, pos_y, pos_z, handle = bloco.values()
            row = ft.DataRow(cells=[
                ft.DataCell(ft.Container(ft.Text(str(tag)), width=140)),
                ft.DataCell(ft.Container(ft.Text(str(tipo)), width=160)),
                ft.DataCell(ft.Container(ft.Text(str(pos_x)), width=120)),
                ft.DataCell(ft.Container(ft.Text(str(pos_y)), width=120)),
                ft.DataCell(ft.Container(ft.Text(str(pos_z)), width=100)),
                ft.DataCell(ft.Container(ft.Text(str(handle)), width=80)),
                ft.DataCell(ft.Row([
                    ft.IconButton(icon=ft.Icons.ZOOM_IN, icon_color="red",
                                  data={'Nome': tag, 'Posicao X': pos_x, 'Posicao Y': pos_y, 'Posicao Z': pos_z},
                                  on_click=zoom_to_block),
                    ft.IconButton(icon=ft.Icons.INFO_OUTLINE, icon_color="blue",
                                  data={'Handle': handle}, on_click=listarPropriedades),
                ])),
            ])
            mydt.rows.append(row)
        page.update()

    mydt = ft.DataTable(
        width=1200, vertical_lines=ft.BorderSide(3, "blue"), horizontal_lines=ft.BorderSide(1, "green"),
        sort_column_index=2, sort_ascending=True, heading_row_color=ft.Colors.BLACK12,
        data_row_color={ft.ControlState.HOVERED: "0x30FF0000"},
        columns=[
            ft.DataColumn(ft.Text('TAG Suporte')), ft.DataColumn(ft.Text('TipoSuporte')),
            ft.DataColumn(ft.Text('X')), ft.DataColumn(ft.Text('Y')), ft.DataColumn(ft.Text('Z')),
            ft.DataColumn(ft.Text('Handle')), ft.DataColumn(ft.Text('actions')),
        ],
        rows=[]
    )

    dt_blocos_valores = ft.DataTable(
        width=1200, vertical_lines=ft.BorderSide(3, "blue"), horizontal_lines=ft.BorderSide(1, "green"),
        sort_column_index=2, sort_ascending=True, heading_row_color=ft.Colors.BLACK12,
        data_row_color={ft.ControlState.HOVERED: "0x30FF0000"},
        columns=[
            ft.DataColumn(ft.Text('Atributo')), ft.DataColumn(ft.Text('Valor')),
            ft.DataColumn(ft.Text('Handle')), ft.DataColumn(ft.Text('actions')),
        ],
        rows=[]
    )

    container_dt_blocos_valores = ft.Container(
        content=ft.Column([
            ft.Container(content=dt_blocos_valores, expand=True)
        ]),
        margin=10,
        padding=10,
        alignment=ft.Alignment.CENTER,
        bgcolor=ft.Colors.GREEN_200,
        height=350,
        border_radius=10,
        visible=False
    )

    def mensagemPop(msgPop, sizePop, colorPop):
        snack_bar = ft.SnackBar(ft.Text(msgPop, size=sizePop), bgcolor=colorPop)
        page.overlay.append(snack_bar)
        snack_bar.open = True
        page.update()
        threading.Thread(target=hide_snack_bar_after_delay, args=(snack_bar,)).start()

    def hide_snack_bar_after_delay(snack_bar):
        time.sleep(3)
        snack_bar.open = False
        page.update()

    def hide_container_after_delay():
        time.sleep(3)
        container_dt_blocos_valores.visible = False
        page.update()

    listarSuportes()

    # debug info to confirm UI rendering and data load
    debug_text = ft.Text(f"Flet version: {getattr(ft, '__version__', 'unknown')} — rows: {len(mydt.rows)}", size=12, color=ft.Colors.BLACK)

    # Create tabs and assign content separately to avoid passing 'content' in Tab constructor
    tab_suportes = ft.Tab(label="Suportes", icon=ft.Icons.SEARCH)
    tab_suportes.content = ft.Container(
        ft.Column([
            debug_text,
            search_option, searchtxt, ft.Button('Buscar', on_click=listarSuportes),
            ft.Container(
                content=ft.Column([
                    ft.Container(content=mydt, expand=True)
                ], scroll=ft.ScrollMode.AUTO),
                expand=True, height=400, bgcolor=ft.Colors.AMBER,
                margin=10, border_radius=10,
            ),
            ft.Row([
                container_dt_blocos_valores
            ]),
        ]),
    )

    tab2 = ft.Tab(label="Tab 2", icon=ft.Icons.SETTINGS)
    tab2.content = ft.Text("Amazing TAB 2 content")

    tab3 = ft.Tab(label="Amazing", icon=ft.Icons.SEARCH)
    tab3.content = ft.Column(controls=[
        ft.Card(
            elevation=30,
            content=ft.Container(
                content=ft.Text("Amazing TAB 1 content", size=50, weight=ft.FontWeight.BOLD),
                border_radius=ft.BorderRadius.all(20), bgcolor=ft.Colors.WHITE24, padding=45,
            )
        )
    ])

    # Create Tabs widget and assign the list of tabs after construction for compatibility
    tabs_widget = ft.Tabs(content=ft.Column(), length=3, selected_index=0, animation_duration=300, expand=3)
    tabs_widget.tabs = [tab_suportes, tab2, tab3]

    # Debug: show tabs info and also add first tab content directly to page to verify rendering
    page.add(ft.Text(f"DEBUG: tabs_widget has {len(tabs_widget.tabs)} tabs", color=ft.Colors.BLUE))
    try:
        labels = [str(t.label) for t in tabs_widget.tabs]
    except Exception:
        labels = [str(type(t)) for t in tabs_widget.tabs]
    page.add(ft.Text("DEBUG: tab labels: " + ", ".join(labels), color=ft.Colors.BLUE))
    page.add(ft.Text(f"DEBUG: mydt rows = {len(mydt.rows)}", color=ft.Colors.BLUE))

    # Also add the raw content of the first tab directly to the page to check rendering
    if tab_suportes.content is not None:
        page.add(ft.Container(content=ft.Column([ft.Text("DEBUG: conteúdo da Tab 'Suportes' abaixo:"), tab_suportes.content]), padding=5))

    page.add(
        ft.Column(
            [tabs_widget]
        )
    )


ft.run(main)
