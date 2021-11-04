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

[PlantUml PowerPoint Add-in 1.1](https://github.com/kmierzeje/plantuml-powerpoint-addin/releases/download/v1.1_fixed/PlantUml.ppam)

## Install

1. Open PowerPoint
2. Go to Developer tab and select `PowerPoint Add-ins`.

  ![obraz](https://user-images.githubusercontent.com/66111032/140281173-6eabfb09-08e0-43e4-bdec-d6393fdcc61b.png) 
  
3. In the popup window select `Add New...` and find `PlantUml.ppam`.
   
   ![obraz](https://user-images.githubusercontent.com/66111032/140281729-5b81f02c-0ec2-4bc0-83b8-d75055f56ad9.png)

4. In the Security Notice that will popup, select `Enable Macros`.
   
   ![obraz](https://user-images.githubusercontent.com/66111032/140282360-c19c1580-3bb6-497d-a296-4e9ad274eea5.png)

## Development and Building

1. Zip the content of `src/PlantUml.pptm` folder to an archive named `PlantUml.pptm`.
2. Use [VbaSync](https://github.com/chelh/VBASync/releases/tag/v2.2.0) to publish source code from `src/vba` to `PlantUml.pptm` file created in first step.

   ![obraz](https://user-images.githubusercontent.com/66111032/138966925-53df51ad-b8d5-4fd5-9e3f-d200cd44de0e.png)

4. Open `PlantUml.pptm` with PowerPoint.
5. Now you can use `View Code` button on `Developer` tab to develop the add-in code.
6. Save As "PowerPoint Add-in" named `PlantUml.ppam`.
7. Don't forget to extract the code from `Plantuml.pptm` with VbaSync back to `src/vba` and issue a Pull Request.


