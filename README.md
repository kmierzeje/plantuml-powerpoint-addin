# plantuml-powerpoint-addin

This add-in allows to embed PlantUml diagrams in PowerPoint presentations.

## Usage

The add-in adds "PlantUml" group to "Insert Tab" with "Insert Diagram" button inside:

![obraz](https://user-images.githubusercontent.com/66111032/138903113-12cc1551-eb24-49d2-a6cd-16e7b01afddf.png)

1. Click the button to insert a new Diagram. A window will popup:

   ![obraz](https://user-images.githubusercontent.com/66111032/138903516-b9b9330d-c9b6-4e63-9249-b7bf92554192.png)

2. Use the "Jar Location" box to enter the location of plantuml.jar.
3. Put your diagram source into the window.
4. Select "Live updates" if you want to see your diagram updated while you type.
5. Close the window after finished editing.
6. If you want to update your diagram, open the context menu and select "Edit PlantUML":

   ![obraz](https://user-images.githubusercontent.com/66111032/138904193-a8c70b1b-b9e8-4f72-8b4d-1e46c42c3af1.png)

## Download

[PlantUml PowerPoint Add-in 1.0](https://github.com/kmierzeje/plantuml-powerpoint-addin/releases/download/v1.0/PlantUml.ppam)

## Install

Opening `PlantUml.ppam` with PowerPoint will install the add-in.

## Build

1. Zip the content of `src/PlantUml.pptm` folder to an archive named `PlantUml.pptm`.
2. Use [VbaSync](https://github.com/chelh/VBASync/releases/tag/v2.2.0) to publish source code from `src/vba` to `PlantUml.pptm` file created in first step.

   ![obraz](https://user-images.githubusercontent.com/66111032/138966925-53df51ad-b8d5-4fd5-9e3f-d200cd44de0e.png)

4. Open `PlantUml.pptm` with PowerPoint.
5. Save As "PowerPoint Add-in" named `PlantUml.ppam`.


