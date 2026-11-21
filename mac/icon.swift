// Disegna l'icona di Parola (1024×1024) e la salva come PNG nel percorso indicato.
import Cocoa

let size: CGFloat = 1024
let out = CommandLine.arguments[1]
let rep = NSBitmapImageRep(bitmapDataPlanes: nil, pixelsWide: Int(size), pixelsHigh: Int(size),
                           bitsPerSample: 8, samplesPerPixel: 4, hasAlpha: true, isPlanar: false,
                           colorSpaceName: .deviceRGB, bytesPerRow: 0, bitsPerPixel: 0)!
NSGraphicsContext.saveGraphicsState()
NSGraphicsContext.current = NSGraphicsContext(bitmapImageRep: rep)

// Riquadro arrotondato con lieve sfumatura (misure della griglia icone macOS: 824px con margine 100)
let rect = NSRect(x: 100, y: 100, width: 824, height: 824)
let shape = NSBezierPath(roundedRect: rect, xRadius: 185, yRadius: 185)
let shadow = NSShadow()
shadow.shadowColor = NSColor.black.withAlphaComponent(0.28)
shadow.shadowOffset = NSSize(width: 0, height: -10)
shadow.shadowBlurRadius = 24
NSGraphicsContext.saveGraphicsState()
shadow.set()
NSColor(srgbRed: 0.10, green: 0.30, blue: 0.70, alpha: 1).setFill()
shape.fill()
NSGraphicsContext.restoreGraphicsState()
let grad = NSGradient(starting: NSColor(srgbRed: 0.27, green: 0.55, blue: 0.96, alpha: 1),
                      ending: NSColor(srgbRed: 0.09, green: 0.28, blue: 0.68, alpha: 1))!
grad.draw(in: shape, angle: -90)

// Lettera W bianca al centro, nel carattere di sistema (SF) molto spesso
let font = NSFont.systemFont(ofSize: 540, weight: .heavy)
let letterShadow = NSShadow()
letterShadow.shadowColor = NSColor.black.withAlphaComponent(0.22)
letterShadow.shadowOffset = NSSize(width: 0, height: -8)
letterShadow.shadowBlurRadius = 18
let text = NSAttributedString(string: "W", attributes: [.font: font, .foregroundColor: NSColor.white])
let ctx = NSGraphicsContext.current!.cgContext
let line = CTLineCreateWithAttributedString(text)
let ink = CTLineGetImageBounds(line, ctx)   // contorno reale della lettera, per centrarla
NSGraphicsContext.saveGraphicsState()
letterShadow.set()
ctx.textPosition = CGPoint(x: 512 - ink.midX, y: 512 - ink.midY)
CTLineDraw(line, ctx)
NSGraphicsContext.restoreGraphicsState()

NSGraphicsContext.restoreGraphicsState()
try! rep.representation(using: .png, properties: [:])!.write(to: URL(fileURLWithPath: out))
