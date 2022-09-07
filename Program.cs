using System;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;

/* 

*/

string docPath = "assets/test.pptx";
string imageName = "assets/line_graph.png";
string replacementImage = "assets/cisco.png";

//sample list of tags
PropertyTag[] propertyTags = new PropertyTag[]{
    new PropertyTag(){tagName = "type" , tagValue = "line-graph"},
    new PropertyTag(){tagName = "id" , tagValue = "123456"},
};

Image.AddImage(docPath, imageName, propertyTags);

// Image.ReplaceImage(docPath, replacementImage);

foreach (var arg in args)
{
    Console.WriteLine(arg);
}


