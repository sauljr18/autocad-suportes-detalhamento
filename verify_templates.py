#!/usr/bin/env python3
"""
verify_templates.py - Verifica se templates DXF estão válidos

Este script verifica os templates DXF convertidos, listando:
- Se o arquivo é válido
- Layouts presentes (Model, PaperSpaces)
- Blocos com atributos encontrados
- Quantidade de atributos por layout

Uso:
    python verify_templates.py [pasta_templates]

    Se pasta não for especificada, usa 'templates' como padrão.
"""

import os
import sys

import ezdxf


def verify_template(template_path):
    """
    Verifica se template DXF é válido e coleta informações.

    Args:
        template_path: Caminho para o arquivo DXF

    Returns:
        Dicionário com informações do template:
        - valid: bool, se o arquivo é válido
        - layouts: lista de dicts com info de cada layout
        - blocks_with_attributes: total de blocos com atributos
        - total_attributes: total de atributos encontrados
        - error: str, mensagem de erro se inválido
    """
    try:
        doc = ezdxf.readfile(template_path)

        info = {
            'valid': True,
            'layouts': [],
            'blocks_with_attributes': 0,
            'total_attributes': 0,
            'error': None
        }

        # Verifica todos os layouts (itera diretamente sobre Layout objects)
        for layout in doc.layouts:
            layout_name = layout.name
            # Pula o Model space - atributos geralmente estão no PaperSpace
            if layout_name == "Model":
                continue
            block_count = 0
            attrib_count = 0
            attrib_tags_found = set()

            # Busca entidades INSERT (block references) com atributos
            for entity in layout:
                if entity.dxftype() == 'INSERT':
                    block_count += 1
                    try:
                        for attrib in entity.attribs:
                            attrib_count += 1
                            # Converte tag para string antes de upper()
                            tag_value = attrib.dxf.tag
                            if tag_value is not None:
                                tag = str(tag_value).upper()
                                attrib_tags_found.add(tag)
                    except Exception as e:
                        # Ignora erros em atributos individuais
                        pass

            if block_count > 0:
                info['layouts'].append({
                    'name': layout_name,
                    'blocks': block_count,
                    'attributes': attrib_count,
                    'tags': sorted(attrib_tags_found)
                })
                info['blocks_with_attributes'] += block_count
                info['total_attributes'] += attrib_count

        return info

    except Exception as e:
        return {
            'valid': False,
            'layouts': [],
            'blocks_with_attributes': 0,
            'total_attributes': 0,
            'error': str(e)
        }


def main():
    """Função principal."""
    template_folder = sys.argv[1] if len(sys.argv) > 1 else "templates"

    if not os.path.isdir(template_folder):
        print(f"Erro: Pasta '{template_folder}' não encontrada!")
        sys.exit(1)

    print(f"Verificando templates em: {template_folder}")
    print("=" * 60)

    dxf_files = [f for f in sorted(os.listdir(template_folder)) if f.endswith('.dxf')]

    if not dxf_files:
        print("Nenhum arquivo .dxf encontrado na pasta.")
        sys.exit(0)

    valid_count = 0
    invalid_count = 0
    total_attribs = 0

    for filename in dxf_files:
        path = os.path.join(template_folder, filename)
        result = verify_template(path)

        print(f"\n📄 {filename}")

        if result['valid']:
            valid_count += 1
            total_attribs += result['total_attributes']
            print(f"   ✅ VÁLIDO")
            print(f"   📐 Layouts com blocos: {len(result['layouts'])}")

            for layout in result['layouts']:
                print(f"      - {layout['name']}: {layout['blocks']} bloco(s), "
                      f"{layout['attributes']} atributo(s)")
                if layout['tags']:
                    print(f"        Tags: {', '.join(layout['tags'])}")

            if result['total_attributes'] == 0:
                print(f"   ⚠️  ATENÇÃO: Nenhum atributo encontrado!")
        else:
            invalid_count += 1
            print(f"   ❌ INVÁLIDO: {result['error']}")

    print("\n" + "=" * 60)
    print(f"RESUMO:")
    print(f"  Arquivos verificados: {len(dxf_files)}")
    print(f"  Válidos: {valid_count}")
    print(f"  Inválidos: {invalid_count}")
    print(f"  Total de atributos: {total_attribs}")
    print("=" * 60)

    if invalid_count > 0:
        sys.exit(1)


if __name__ == "__main__":
    main()
