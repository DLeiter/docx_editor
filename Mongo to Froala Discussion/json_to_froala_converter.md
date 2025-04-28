# Converting DOCX JSON to Froala-Compatible HTML

## Overview

The JSON structure exported from DOCX files is a direct representation of the XML with namespaces and complex nesting. To use this with Froala Editor in Angular, we need a conversion strategy that transforms this structure into clean HTML while preserving the essential formatting.

## Conversion Strategy

There are two main approaches:

1. **Direct Transformation**: Write a custom parser that directly converts DOCX JSON to HTML
2. **Intermediate Libraries**: Use established libraries to handle the conversion

The second approach is recommended for reliability and maintenance.

## Recommended Solution

### Backend (.NET) Approach

```csharp
// Controller endpoint to convert DOCX JSON to HTML
[HttpPost("convert-to-html")]
public async Task<IActionResult> ConvertDocxJsonToHtml([FromBody] JObject docxJson)
{
    try
    {
        // Convert JSON back to XML first
        string xmlContent = JsonToXmlConverter.Convert(docxJson);
        
        // Create a temporary DOCX file from the XML
        string tempDocxPath = Path.GetTempFileName() + ".docx";
        
        // Create the docx package structure
        using (var package = Package.Open(tempDocxPath, FileMode.Create))
        {
            // Add document.xml to the package
            Uri documentUri = new Uri("/word/document.xml", UriKind.Relative);
            using (var part = package.CreatePart(documentUri, "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"))
            {
                using (var stream = part.GetStream())
                using (var writer = new StreamWriter(stream))
                {
                    writer.Write(xmlContent);
                }
            }
            
            // Add other required parts (minimal set)
            // (content types, rels, etc.)
        }
        
        // Now use a library like DocumentFormat.OpenXml or Aspose.Words to convert to HTML
        string html;
        using (var doc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(tempDocxPath, false))
        {
            // Convert to HTML using a library like HtmlConverter
            html = HtmlConverter.ConvertToHtml(doc);
        }
        
        // Clean up temp file
        if (File.Exists(tempDocxPath))
            File.Delete(tempDocxPath);
        
        return Ok(new { html });
    }
    catch (Exception ex)
    {
        return BadRequest(new { error = ex.Message });
    }
}

// Utility class for JSON to XML conversion
public static class JsonToXmlConverter
{
    public static string Convert(JObject json)
    {
        // Convert JSON back to XML
        XmlDocument xmlDoc = new XmlDocument();
        
        // Create the XML declaration
        XmlDeclaration declaration = xmlDoc.CreateXmlDeclaration("1.0", "UTF-8", null);
        xmlDoc.AppendChild(declaration);
        
        // Create the document element with all namespaces
        XmlElement documentElement = xmlDoc.CreateElement("w", "document", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        
        // Add namespaces from JSON
        var documentJson = json["w:document"];
        foreach (JProperty prop in documentJson.Properties())
        {
            if (prop.Name.StartsWith("@xmlns:"))
            {
                string prefix = prop.Name.Substring(7); // Remove '@xmlns:'
                documentElement.SetAttribute("xmlns:" + prefix, prop.Value.ToString());
            }
            else if (prop.Name == "@mc:Ignorable")
            {
                documentElement.SetAttribute("mc:Ignorable", "http://schemas.openxmlformats.org/markup-compatibility/2006", prop.Value.ToString());
            }
        }
        
        xmlDoc.AppendChild(documentElement);
        
        // Process body and its children
        ProcessJsonToXml(documentJson["w:body"], documentElement, xmlDoc);
        
        return xmlDoc.OuterXml;
    }
    
    private static void ProcessJsonToXml(JToken jsonNode, XmlElement parentElement, XmlDocument xmlDoc)
    {
        if (jsonNode == null) return;
        
        foreach (JProperty prop in jsonNode.Properties())
        {
            if (prop.Name.StartsWith("@")) // Attribute
            {
                string attributeName = prop.Name.Substring(1);
                string[] parts = attributeName.Split(':');
                if (parts.Length > 1)
                {
                    string prefix = parts[0];
                    string localName = parts[1];
                    string ns = GetNamespaceUri(prefix);
                    parentElement.SetAttribute(localName, ns, prop.Value.ToString());
                }
                else
                {
                    parentElement.SetAttribute(attributeName, prop.Value.ToString());
                }
            }
            else // Child element
            {
                string[] nameParts = prop.Name.Split(':');
                string prefix = nameParts.Length > 1 ? nameParts[0] : string.Empty;
                string localName = nameParts.Length > 1 ? nameParts[1] : nameParts[0];
                string ns = GetNamespaceUri(prefix);
                
                if (prop.Value.Type == JTokenType.Object)
                {
                    XmlElement childElement = xmlDoc.CreateElement(prefix, localName, ns);
                    parentElement.AppendChild(childElement);
                    ProcessJsonToXml(prop.Value, childElement, xmlDoc);
                }
                else if (prop.Value.Type == JTokenType.Array)
                {
                    foreach (var item in prop.Value)
                    {
                        XmlElement childElement = xmlDoc.CreateElement(prefix, localName, ns);
                        parentElement.AppendChild(childElement);
                        ProcessJsonToXml(item, childElement, xmlDoc);
                    }
                }
                else // Simple value
                {
                    XmlElement childElement = xmlDoc.CreateElement(prefix, localName, ns);
                    childElement.InnerText = prop.Value.ToString();
                    parentElement.AppendChild(childElement);
                }
            }
        }
    }
    
    private static string GetNamespaceUri(string prefix)
    {
        // Map common prefixes to their namespaces - should be expanded based on your document
        switch (prefix)
        {
            case "w": return "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            case "r": return "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            // Add other namespace mappings as needed
            default: return string.Empty;
        }
    }
}
```

### Frontend (Angular) Approach

For a frontend-only solution, you can use Mammoth.js which is specifically designed for converting DOCX to HTML:

```typescript
import { Component } from '@angular/core';
import * as mammoth from 'mammoth';
import { HttpClient } from '@angular/common/http';

@Component({
  selector: 'app-document-converter',
  template: `
    <div>
      <h2>DOCX JSON to HTML Converter</h2>
      
      <div *ngIf="isConverting" class="spinner">Converting...</div>
      
      <div *ngIf="!isConverting">
        <div class="form-group">
          <label>Upload DOCX File</label>
          <input type="file" (change)="onFileSelected($event)" accept=".docx">
        </div>
        
        <div *ngIf="htmlContent" class="preview">
          <h3>HTML Preview</h3>
          <div [froalaEditor]="froalaOptions" [(froalaModel)]="htmlContent"></div>
          
          <button class="btn btn-primary" (click)="saveToDatabase()">Save to Database</button>
        </div>
      </div>
    </div>
  `
})
export class DocumentConverterComponent {
  htmlContent: string = '';
  isConverting: boolean = false;
  docxJson: any = null;
  
  froalaOptions: any = {
    // Froala options
  };
  
  constructor(private http: HttpClient) {}
  
  async onFileSelected(event: any) {
    const file = event.target.files[0];
    if (!file) return;
    
    this.isConverting = true;
    
    try {
      // If using Mammoth.js directly (frontend only approach)
      const arrayBuffer = await file.arrayBuffer();
      const result = await mammoth.convertToHtml({ arrayBuffer });
      this.htmlContent = result.value;
      
      // Alternative: Send to backend for conversion
      /*
      const formData = new FormData();
      formData.append('file', file);
      
      this.http.post<any>('/api/documents/convert-to-html', formData)
        .subscribe(
          response => {
            this.htmlContent = response.html;
            this.isConverting = false;
          },
          error => {
            console.error('Error converting file', error);
            this.isConverting = false;
          }
        );
      */
    } catch (error) {
      console.error('Error converting file', error);
    } finally {
      this.isConverting = false;
    }
  }
  
  saveToDatabase() {
    if (!this.htmlContent) return;
    
    // Create the document structure for MongoDB
    const documentData = {
      title: 'Imported Document',
      htmlCache: this.htmlContent,
      structuredContent: this.extractStructuredContent(this.htmlContent),
      metadata: {
        createdDate: new Date(),
        lastModified: new Date()
      }
    };
    
    // Send to backend for saving
    this.http.post<any>('/api/documents', documentData)
      .subscribe(
        response => {
          console.log('Document saved', response);
          // Navigate to editor or show success message
        },
        error => {
          console.error('Error saving document', error);
        }
      );
  }
  
  extractStructuredContent(html: string): any {
    // Parse HTML and extract structured content for MongoDB
    const parser = new DOMParser();
    const doc = parser.parseFromString(html, 'text/html');
    
    const structuredContent = {
      paragraphs: []
    };
    
    // Process paragraphs
    doc.querySelectorAll('p, h1, h2, h3, h4, h5, h6').forEach((element, index) => {
      const isHeading = element.tagName.toLowerCase().startsWith('h');
      const headingLevel = isHeading ? parseInt(element.tagName.substring(1)) : null;
      
      structuredContent.paragraphs.push({
        paragraphId: `p-${index}`,
        text: element.textContent,
        html: element.innerHTML,
        style: isHeading ? `Heading${headingLevel}` : 'Normal',
        formatting: {
          bold: element.querySelector('strong, b') !== null,
          italic: element.querySelector('em, i') !== null,
          underline: element.querySelector('u') !== null
        },
        lastModified: new Date()
      });
    });
    
    return structuredContent;
  }
}
```

## Library-Based Approach (Recommended)

For production use, I recommend a library-based approach:

### Backend Libraries

1. **Aspose.Words** (Commercial, comprehensive)
```csharp
// Convert JSON to DOCX and then to HTML
public string ConvertJsonToHtml(JObject docxJson)
{
    // Convert JSON to XML
    string xml = JsonToXmlConverter.Convert(docxJson);
    
    // Load the XML into a MemoryStream
    using (MemoryStream xmlStream = new MemoryStream(Encoding.UTF8.GetBytes(xml)))
    {
        // Create a minimal DOCX file
        MemoryStream docxStream = CreateMinimalDocx(xmlStream);
        
        // Use Aspose.Words to convert to HTML
        using (var doc = new Aspose.Words.Document(docxStream))
        {
            var saveOptions = new Aspose.Words.Saving.HtmlSaveOptions
            {
                ExportImagesAsBase64 = true,
                ExportFontResources = true,
                CssStyleSheetType = Aspose.Words.Saving.CssStyleSheetType.Embedded
            };
            
            MemoryStream htmlStream = new MemoryStream();
            doc.Save(htmlStream, saveOptions);
            htmlStream.Position = 0;
            
            using (StreamReader reader = new StreamReader(htmlStream))
            {
                return reader.ReadToEnd();
            }
        }
    }
}
```

2. **OpenXML SDK + HtmlConverter** (Free, more work)
```csharp
// Using OpenXML SDK with the HtmlConverter package
using DocumentFormat.OpenXml.Packaging;
using HtmlAgilityPack;

public string ConvertDocxJsonToHtml(JObject docxJson)
{
    // Similar approach as above, but using OpenXML SDK
    // with a third-party HTML converter like OpenXmlPowerTools
}
```

### Frontend Libraries

1. **Mammoth.js** (Recommended)
```typescript
import * as mammoth from 'mammoth';

async function convertDocxToHtml(file: File): Promise<string> {
    const arrayBuffer = await file.arrayBuffer();
    const result = await mammoth.convertToHtml({ arrayBuffer });
    return result.value;
}
```

2. **docx-preview** (Alternative)
```typescript
import { renderAsync } from 'docx-preview';

async function renderDocxInBrowser(file: File, container: HTMLElement): Promise<void> {
    const arrayBuffer = await file.arrayBuffer();
    await renderAsync(arrayBuffer, container);
    // Extract HTML from container if needed
}
```

## Custom Mapping (Most Control)

For the most control over the conversion process, you can implement a custom mapping between DOCX elements and HTML:

```typescript
interface DocxParagraph {
  style?: string;
  text?: string;
  runs?: DocxRun[];
}

interface DocxRun {
  text: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
}

function convertStructuredJsonToHtml(docxJson: any): string {
  const html = ['<div class="document">'];
  
  // Extract paragraphs from the complex JSON structure
  const paragraphs = extractParagraphs(docxJson);
  
  for (const para of paragraphs) {
    const style = para.style || 'Normal';
    
    // Convert paragraph style to HTML element
    let element = 'p';
    let className = '';
    
    if (style.startsWith('Heading')) {
      const level = parseInt(style.replace('Heading', ''));
      if (level >= 1 && level <= 6) {
        element = `h${level}`;
      }
    } else {
      className = ` class="${style.toLowerCase()}"`;
    }
    
    html.push(`<${element}${className}>`);
    
    // Process runs (text with formatting)
    if (para.runs && para.runs.length > 0) {
      for (const run of para.runs) {
        let text = run.text || '';
        
        // Apply formatting
        if (run.bold) text = `<strong>${text}</strong>`;
        if (run.italic) text = `<em>${text}</em>`;
        if (run.underline) text = `<u>${text}</u>`;
        
        html.push(text);
      }
    } else if (para.text) {
      html.push(para.text);
    }
    
    html.push(`</${element}>`);
  }
  
  html.push('</div>');
  return html.join('');
}

// Extract paragraphs from the nested DOCX JSON structure
function extractParagraphs(docxJson: any): DocxParagraph[] {
  const paragraphs: DocxParagraph[] = [];
  
  // Navigate through the nested structure
  try {
    const body = docxJson["w:document"]["w:body"];
    const paraElements = body["w:p"];
    
    if (Array.isArray(paraElements)) {
      // Multiple paragraphs
      for (const para of paraElements) {
        paragraphs.push(processParagraph(para));
      }
    } else if (paraElements) {
      // Single paragraph
      paragraphs.push(processParagraph(paraElements));
    }
  } catch (e) {
    console.error('Error extracting paragraphs', e);
  }
  
  return paragraphs;
}

function processParagraph(para: any): DocxParagraph {
  const result: DocxParagraph = { runs: [] };
  
  // Extract paragraph style
  try {
    const pPr = para["w:pPr"];
    if (pPr && pPr["w:pStyle"] && pPr["w:pStyle"]["@w:val"]) {
      result.style = pPr["w:pStyle"]["@w:val"];
    }
  } catch (e) {}
  
  // Process text runs
  try {
    const runs = para["w:r"];
    if (Array.isArray(runs)) {
      for (const run of runs) {
        result.runs.push(processRun(run));
      }
    } else if (runs) {
      result.runs.push(processRun(runs));
    }
  } catch (e) {}
  
  return result;
}

function processRun(run: any): DocxRun {
  const result: DocxRun = { text: '' };
  
  // Extract text
  try {
    const text = run["w:t"];
    if (text) {
      result.text = typeof text === 'string' ? text : (text["#text"] || '');
    }
  } catch (e) {}
  
  // Extract formatting
  try {
    const rPr = run["w:rPr"];
    if (rPr) {
      result.bold = !!rPr["w:b"];
      result.italic = !!rPr["w:i"];
      result.underline = !!rPr["w:u"];
    }
  } catch (e) {}
  
  return result;
}
```

## Recommended Workflow

1. **Backend Conversion**: Use Aspose.Words or a similar library on your .NET backend
2. **Frontend Conversion**: Use Mammoth.js for client-side conversion
3. **Store Both Formats**: Save both the structured content and HTML in MongoDB
4. **Editor Integration**: Use the HTML with Froala, and convert back to structured data on save

This approach gives you the best of both worlds: Froala gets clean HTML for editing, while you maintain structured data for querying and manipulation.

## Example Pipeline

1. User uploads DOCX
2. Backend extracts JSON structure
3. JSON is converted to clean HTML
4. HTML is loaded into Froala
5. User edits content
6. On save, HTML is parsed back to structured content
7. Both HTML and structured content are stored in MongoDB

This pipeline maintains the content in both formats, ensuring optimal editing experience while preserving the document structure for advanced features.
