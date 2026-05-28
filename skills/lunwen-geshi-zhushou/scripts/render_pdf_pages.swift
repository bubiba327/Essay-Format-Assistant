import AppKit
import Foundation
import PDFKit

func fail(_ message: String) -> Never {
    FileHandle.standardError.write((message + "\n").data(using: .utf8)!)
    exit(1)
}

let args = CommandLine.arguments
guard args.count == 5 else {
    fail("Usage: render_pdf_pages.swift <input.pdf> <output_dir> <max_width_px> <max_height_px>")
}

let pdfURL = URL(fileURLWithPath: args[1])
let outputURL = URL(fileURLWithPath: args[2], isDirectory: true)
guard let maxWidth = Double(args[3]), let maxHeight = Double(args[4]), maxWidth > 0, maxHeight > 0 else {
    fail("max_width_px and max_height_px must be positive numbers.")
}

guard let document = PDFDocument(url: pdfURL) else {
    fail("Cannot open PDF: \(pdfURL.path)")
}

try FileManager.default.createDirectory(at: outputURL, withIntermediateDirectories: true)

for index in 0..<document.pageCount {
    guard let page = document.page(at: index) else {
        fail("Cannot read PDF page \(index + 1)")
    }
    let bounds = page.bounds(for: .mediaBox)
    let scale = min(maxWidth / bounds.width, maxHeight / bounds.height)
    let pixelSize = NSSize(width: bounds.width * scale, height: bounds.height * scale)
    let image = page.thumbnail(of: pixelSize, for: .mediaBox)

    guard
        let tiff = image.tiffRepresentation,
        let bitmap = NSBitmapImageRep(data: tiff),
        let png = bitmap.representation(using: .png, properties: [:])
    else {
        fail("Cannot encode page \(index + 1) as PNG")
    }

    let pageURL = outputURL.appendingPathComponent("page-\(index + 1).png")
    try png.write(to: pageURL)
}

print("Rendered \(document.pageCount) page(s) to \(outputURL.path)")
