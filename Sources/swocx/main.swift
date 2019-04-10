import AppKit


let data = NSData(contentsOfFile: "testdoc.docx")
let str = try! NSAttributedString(data: data! as Data, options: [.documentType: NSAttributedString.DocumentType.officeOpenXML, .characterEncoding: String.Encoding.utf8.rawValue], documentAttributes: nil)
