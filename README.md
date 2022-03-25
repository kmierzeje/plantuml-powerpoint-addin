# plantuml-powerpoint-addin

This add-in allows to embed PlantUml diagrams in PowerPoint presentations.

## Usage

The add-in adds `PlantUML` group to `Insert Tab` with `Insert Diagram` button inside:

![obraz](https://user-images.githubusercontent.com/66111032/138903113-12cc1551-eb24-49d2-a6cd-16e7b01afddf.png)

1. Click the button to insert a new Diagram. A window will popup:

   ![obraz](https://user-images.githubusercontent.com/66111032/139954968-5902aad0-9a7d-43ed-89d0-b302d3d0248d.png)

2. When using for the first time, use the `Jar location` box to enter the location of `plantuml.jar`.
3. Edit your diagram.
4. Close the window after finished editing.
5. If you want to update your diagram, open the context menu and select `Edit PlantUML`:

   ![obraz](https://user-images.githubusercontent.com/66111032/138904193-a8c70b1b-b9e8-4f72-8b4d-1e46c42c3af1.png)
   
6. The diagram editor window will popup again.

## Download

[PlantUml PowerPoint Add-in 1.3](https://github.com/kmierzeje/plantuml-powerpoint-addin/releases/download/v1.3/PlantUml.ppam)

## Install

1. Make sure you have Java Runtime Environment installed and `java` executable location in `PATH`.
2. Open PowerPoint.
3. Go to Developer tab and select `PowerPoint Add-ins`.

  ![obraz](https://user-images.githubusercontent.com/66111032/140281173-6eabfb09-08e0-43e4-bdec-d6393fdcc61b.png) 
  
3. In the popup window select `Add New...` and find `PlantUml.ppam`.
   
   ![obraz](https://user-images.githubusercontent.com/66111032/140281729-5b81f02c-0ec2-4bc0-83b8-d75055f56ad9.png)

4. In the Security Notice that will popup, select `Enable Macros`.
   
   ![obraz](https://user-images.githubusercontent.com/66111032/140282360-c19c1580-3bb6-497d-a296-4e9ad274eea5.png)

## Development and Building

### Prerequisites

- [VBA Sync Tool](https://github.com/chelh/VBASync)
- [Zip](http://infozip.sourceforge.net/Zip.html)

### Building

Run `build.bat path/to/target/PlantUml.ppam` in `src` directory.

### Development

1. Run `build.bat path/to/target/PlantUml.pptm` in `src` directory.
2. Open `path/to/target/PlantUml.pptm` with PowerPoint.
3. Now you can use `View Code` button on `Developer` tab to develop the add-in code.
4. In VBA Editor use `File/Export File...` menu to save updated files in `src/vba`. You can also run `VBA Sync Tool` GUI to export your changes.
5. Don't forget to issue a Pull Request.
