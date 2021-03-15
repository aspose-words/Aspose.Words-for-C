#include "hidpisupport.h"

#include <QApplication>
#include <QPixmap>
#include <QScreen>

namespace HiDpiSupport {

float getImageScaleFactor()
{
    const float standartDpi = 96.0f;
    const float screenDpi = QApplication::screens().at(0)->logicalDotsPerInch();
    return screenDpi / standartDpi;
}

QPixmap scaleImage(const QPixmap& image)
{
    return image.scaled(image.size() * getImageScaleFactor(),
                        Qt::IgnoreAspectRatio,
                        Qt::SmoothTransformation);
}

} // namespace HiDpiSupport
