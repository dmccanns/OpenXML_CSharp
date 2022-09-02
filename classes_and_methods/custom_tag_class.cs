using System;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
class CustomTag{
    public static void AddCustomTagToPicture(string propertyName, string propertyValue, SlidePart slidePart , Picture picture){
        
    //creates the tag.xml file
    TagList tagList = new TagList();
    Tag tag = new Tag() { Name = propertyName, Val = propertyValue };
    tagList.Append(tag);

    //TODO: create id programatically
    var userTagsPart = slidePart.AddNewPart<UserDefinedTagsPart>("rId4");
    userTagsPart.TagList = tagList;

    //create the reference to the tags and append it in the picture
    var propertyElement = new ApplicationNonVisualDrawingProperties();
    var dataListElement = new CustomerDataList();
    var refTag = new CustomerDataTags(){
        Id = "rId4"
    };
    dataListElement.Append(refTag);
    propertyElement.Append(dataListElement);
    picture.Append(propertyElement);
    }
}