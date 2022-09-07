/*
Top level program
    To view XML I used the OOXML viewer extension in VSCode by Matthew Yuen. 
    Extension ID: yuenm18.ooxml-viewer
*/
string docPath = "assets/test.pptx";
string imagePath = "assets/line_graph.png";
string replacementImage = "assets/cisco.png";

//sample properties - these properties are needed for running queries
CustomProperties properties = new CustomProperties
{
    type = "line-graph",
    project = 123456,
    workspace = 123456,
    bookmark = 123456,
};

/*
Function to add image to specified doc w/ properties. Called by Electron GUI
*/
Image.AddImage(docPath, imagePath, properties);

/*
Function to replace images. Currently one function but should be split into two functions.

    First function should return list of properties so query can be run, second should get list
of properties, check them against arguments + image paths, then use those images to replace
the existing ones.
*/
// Image.ReplaceImage(docPath, replacementImage);


