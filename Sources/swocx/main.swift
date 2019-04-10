import AppKit





func readDocx(_ infile: String) -> NSAttributedString {
    let data = NSData(contentsOfFile: infile)
    let str = try! NSAttributedString(data: data! as Data, options: [.documentType: NSAttributedString.DocumentType.officeOpenXML, .characterEncoding: String.Encoding.utf8.rawValue], documentAttributes: nil)
    return str
}

func readMutableDocx(_ infile: String) -> NSMutableAttributedString {
    let data = NSData(contentsOfFile: infile)
    let str = try! NSMutableAttributedString(data: data! as Data, options: [.documentType: NSAttributedString.DocumentType.officeOpenXML, .characterEncoding: String.Encoding.utf8.rawValue], documentAttributes: nil)
    return str
}



func listDocsInCurrentDirectory() throws -> [String] {
    let fileManager = FileManager()
    let files = try! fileManager.contentsOfDirectory(atPath: ".") 
    let docs = files.filter({ $0.hasSuffix(".docx") }).sorted(by: <)
    return docs
}

func mergeDocs() -> NSMutableAttributedString {
    let docs = try! listDocsInCurrentDirectory()
    let first = readMutableDocx(docs[0])
    let rest = docs[1...]
    let lineBreaks = NSAttributedString(string: "\n\n")
    for docToAdd in rest {
        first.append(lineBreaks)
        first.append(readDocx(docToAdd))
    }
    return first
}

func toDocx(_ str: NSAttributedString) -> Data {
    let range = NSMakeRange(0, str.length)
    let out = try! str.data(from: range, documentAttributes: [.documentType: NSAttributedString.DocumentType.officeOpenXML])
    return out
}

let docdata = mergeDocs()

let docx = toDocx(docdata)

let dest = URL(fileURLWithPath: "testSwiftDocxOutput.docx")
try! docx.write(to: dest)
