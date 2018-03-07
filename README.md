# Excel 'Graphics Layer'
Api to simulate OpenGL in Excel. 

- Just a basic functionality is implemented, enough to render something on screen 
- Support for multibuffering. I did not notice any improvements in using multiple sheets as buffers
- Easy configuration of the display surface
- Shader support


Limitations
- Excel has a limit on the number of styles it can have, in our case it means the number of different colored cells, so you can get into a situation where you get the error message "Too many different cell formats"
- No support for depth buffer on the "Mesh" drawing state, need to find a way to draw the lines and keep track of the z coordinate
- The triangles are rasterized only if they are in CW winding.
- Speed :), on an i7-3770, Excel 2016, I get at most 40k/s pixel fill rate. Keep in mind that the fps seems to be capped at 30 due to the need for worksheet updates.


Future. Maybe. Remove the matrix stack, and use a more modern api, let the user handle the matrices in the vertex shader.

![](https://raw.githubusercontent.com/MRazvan/egl/master/screenshots/Triangle.png)
![](https://raw.githubusercontent.com/MRazvan/egl/master/screenshots/AnotherCube.png)
![](https://raw.githubusercontent.com/MRazvan/egl/master/screenshots/Cube.png)




Youtube
https://www.youtube.com/watch?v=2AvyvVFONp8
