# Automoção sobre Sondagem a Percurssão
By [Artillis Prado](https://github.com/artillisprado)

## Getting Started / Introdução

### Como encontrar dll, para o civil 3D
  - <p>caminho: C:\Program Files\Autodesk\AutoCAD 2023\C3D </p> 
  - <p>Nomes das dll: 
        accoremgd.dll,
        AcCui.dll,
        acdbmgd.dll,
        acdbmgdbrep.dll,
        acmgd.dll,
        AdWindows.dll,
        AecBaseMgd.dll,
        AeccDbMgd.dll,
        AeccUiWindows.dll,
        AeccUiWindows.Resources.dll,
        AeccXUiLand.dll,
        AecPropDataMgd.dll,
        AecPropertyManagement.dll,
        Autodesk.AECC.Interop.Survey.dll,
        Autodesk.AECC.Interop.UiSurvey.dll,
        Autodesk.AutoCAD.Interop.dll,
        Autodesk.Map.Platform.dll,
        ManagedMapApi.dll,
        OSGeo.MapGuide.Foundation.dll,
        OSGeo.MapGuide.Geometry.dll,
        OSGeo.MapGuide.PlatformBase.dll</p>
  - <p>Salvar arquivo em uma pasta "extensões civil 3d"</p>

  ### Adicionar dll no Visual Studio
    - Criando um projeto (Novo -> projeto)
      └── Criar novo projeto -> Biblioteca de Classes -> aplicativo de console
    - Gerenciador de Soluções:
      └── nome_projeto -> Referências -> extensões civil 3d -> Adicione todas as dll
