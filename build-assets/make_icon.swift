import AppKit

let outputDir = URL(fileURLWithPath: FileManager.default.currentDirectoryPath)
  .appendingPathComponent("icon.iconset", isDirectory: true)

try? FileManager.default.removeItem(at: outputDir)
try FileManager.default.createDirectory(at: outputDir, withIntermediateDirectories: true)

let sizes: [(String, CGFloat)] = [
  ("icon_16x16.png", 16), ("icon_16x16@2x.png", 32),
  ("icon_32x32.png", 32), ("icon_32x32@2x.png", 64),
  ("icon_128x128.png", 128), ("icon_128x128@2x.png", 256),
  ("icon_256x256.png", 256), ("icon_256x256@2x.png", 512),
  ("icon_512x512.png", 512), ("icon_512x512@2x.png", 1024),
]

func makeImage(size: CGFloat) -> NSImage {
  let image = NSImage(size: NSSize(width: size, height: size))
  image.lockFocus()

  let rect = NSRect(x: 0, y: 0, width: size, height: size)
  let background = NSBezierPath(roundedRect: rect.insetBy(dx: size * 0.04, dy: size * 0.04), xRadius: size * 0.22, yRadius: size * 0.22)
  let gradient = NSGradient(colors: [
    NSColor(calibratedRed: 0.08, green: 0.52, blue: 1.00, alpha: 1.0),
    NSColor(calibratedRed: 0.13, green: 0.69, blue: 0.96, alpha: 1.0),
  ])!
  gradient.draw(in: background, angle: -45)

  let inset = rect.insetBy(dx: size * 0.12, dy: size * 0.12)
  let inner = NSBezierPath(roundedRect: inset, xRadius: size * 0.16, yRadius: size * 0.16)
  NSColor(calibratedWhite: 1.0, alpha: 0.18).setFill()
  inner.fill()

  let paragraph = NSMutableParagraphStyle()
  paragraph.alignment = .center

  let title = "hwp"
  let subtitle = "AI OFFICE"

  let titleAttributes: [NSAttributedString.Key: Any] = [
    .font: NSFont.systemFont(ofSize: size * 0.30, weight: .bold),
    .foregroundColor: NSColor.white,
    .paragraphStyle: paragraph,
  ]
  let subtitleAttributes: [NSAttributedString.Key: Any] = [
    .font: NSFont.systemFont(ofSize: size * 0.075, weight: .semibold),
    .foregroundColor: NSColor(calibratedWhite: 1.0, alpha: 0.84),
    .paragraphStyle: paragraph,
  ]

  let titleRect = NSRect(x: size * 0.12, y: size * 0.34, width: size * 0.76, height: size * 0.24)
  let subtitleRect = NSRect(x: size * 0.16, y: size * 0.20, width: size * 0.68, height: size * 0.10)
  title.draw(in: titleRect, withAttributes: titleAttributes)
  subtitle.draw(in: subtitleRect, withAttributes: subtitleAttributes)

  let spark = NSBezierPath()
  spark.lineWidth = max(2, size * 0.018)
  spark.move(to: NSPoint(x: size * 0.26, y: size * 0.72))
  spark.line(to: NSPoint(x: size * 0.74, y: size * 0.72))
  spark.move(to: NSPoint(x: size * 0.34, y: size * 0.78))
  spark.line(to: NSPoint(x: size * 0.66, y: size * 0.78))
  NSColor(calibratedWhite: 1.0, alpha: 0.92).setStroke()
  spark.stroke()

  image.unlockFocus()
  return image
}

for (name, size) in sizes {
  let image = makeImage(size: size)
  let data = image.tiffRepresentation!
  let rep = NSBitmapImageRep(data: data)!
  let png = rep.representation(using: .png, properties: [:])!
  try png.write(to: outputDir.appendingPathComponent(name))
}
