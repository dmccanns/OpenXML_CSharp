using System;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;

string docPath = "assets/test-2.pptx";
string imageName = "assets/line_graph.png";
string replacementImage = "assets/cisco.png";
// Console.WriteLine(NumberOfSlides.RetrieveNumberOfSlides(docPath, true));

// Image.AddImage(docPath, imageName, "type", "test value");

Image.ReplaceImage(docPath, replacementImage);

//now run the program which replaces an image based on its tag
