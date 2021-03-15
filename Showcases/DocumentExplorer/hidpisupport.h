#ifndef HIDPISUPPORT_H
#define HIDPISUPPORT_H

class QPixmap;

namespace HiDpiSupport {

/// Get image scale factor (1.0 is for 96 dpi).
float getImageScaleFactor();

/// Scale image to screen dpi.
QPixmap scaleImage(const QPixmap& image);

} // namespace HiDpiSupport

#endif // HIDPISUPPORT_H
