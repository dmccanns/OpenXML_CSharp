using System;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;

string docPath = "assets/test.pptx";
string imageName = "assets/line_graph.png";
Console.WriteLine(NumberOfSlides.RetrieveNumberOfSlides(docPath, true));

Image.AddImage(docPath, imageName);

//Image.ReplaceImage(docPath, "SHAPETAG");

//now run the program which replaces an image based on its tag
