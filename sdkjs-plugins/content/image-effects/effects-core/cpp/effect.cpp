#include "effect.h"

bool ChangeBrightnessEffect::Apply(Image *img)
{
    if(brightness_ == 0) return false;
    int h = img->height();
    int w = img->width();
    PixelHSV hsv;

    for(int i = 0; i < h; i++)
        for(int j = 0; j < w; j++)
        {
            hsv.fromRGB(img->at(i, j).RGB());
            hsv.V += (double)brightness_ / 100;

            // проверяем границы
            if(hsv.V > 1.0) hsv.V = 1;
            if(hsv.V < 0.0) hsv.V = 0;

            img->at(i, j).setRGB(hsv.RGB());
        }
    return true;
}
ChangeBrightnessEffect::ChangeBrightnessEffect() : brightness_(0)
{
}
ChangeBrightnessEffect::ChangeBrightnessEffect(int brightness) : brightness_(brightness)
{
}
int ChangeBrightnessEffect::brightness()
{
    return brightness_;
}
void ChangeBrightnessEffect::setBrightness(int brightness)
{
    brightness_ = brightness;
}

bool ChangeSaturationEffect::Apply(Image *img)
{
    if(saturation_ == 0) return false;

    int h = img->height();
    int w = img->width();
    PixelHSV hsv;

    for(int i = 0; i < h; i++)
        for(int j = 0; j < w; j++)
        {
            hsv.fromRGB(img->at(i, j).RGB());
            hsv.S += (double)saturation_ / 100;

            // проверяем границы
            if(hsv.S > 1.0) hsv.S = 1;
            if(hsv.S < 0.0) hsv.S = 0;

            img->at(i, j).setRGB(hsv.RGB());
        }
    return true;
}
ChangeSaturationEffect::ChangeSaturationEffect() : saturation_(0)
{
}
ChangeSaturationEffect::ChangeSaturationEffect(int saturation) : saturation_(saturation)
{
}
int ChangeSaturationEffect::saturation()
{
    return saturation_;
}
void ChangeSaturationEffect::setSaturation(int saturation)
{
    saturation_ = saturation;
}

bool ChangeHueEffect::Apply(Image *img)
{
    if(hue_ == 0) return false;

    int h = img->height();
    int w = img->width();
    PixelHSV hsv;

    for(int i = 0; i < h; i++)
        for(int j = 0; j < w; j++)
        {
            hsv.fromRGB(img->at(i, j).RGB());
            hsv.H += hue_;

            // проверяем границы
            if(hsv.H > 360) hsv.H = hsv.H - 360;
            if(hsv.H < 0) hsv.H = hsv.H + 360;

            img->at(i, j).setRGB(hsv.RGB());
        }
    return true;
}
ChangeHueEffect::ChangeHueEffect() : hue_(0)
{
}
ChangeHueEffect::ChangeHueEffect(int hue) : hue_(hue)
{
}
int ChangeHueEffect::hue()
{
    return hue_;
}
void ChangeHueEffect::setHue(int hue)
{
    hue_ = hue;
}

bool ChangeTransparencyEffect::Apply(Image *img)
{
    if(transparency_ == 0) return false;

    int h = img->height();
    int w = img->width();
    int t = transparency_ * 255 / 100;
    int a = 0;
    int new_a = 0;

    for(int i = 0; i < h; i++)
        for(int j = 0; j < w; j++)
        {
            a = img->at(i, j).A();
            new_a = a - t;

            // проверяем границы
            if(new_a > 255) a = 255;
            else if(new_a < 0) a = 0;

            img->at(i, j).setA((unsigned char)new_a);
        }
    return true;
}
ChangeTransparencyEffect::ChangeTransparencyEffect() : transparency_(0)
{
}
ChangeTransparencyEffect::ChangeTransparencyEffect(int transparency) : transparency_(transparency)
{
}
int ChangeTransparencyEffect::transparency()
{
    return transparency_;
}
void ChangeTransparencyEffect::setTransparency(int transparency)
{
    transparency_ = transparency;
}

bool ChangeLightingEffect::Apply(Image *img)
{
    if(lighting_ == 0) return false;

    int h = img->height();
    int w = img->width();
    double l = (double)lighting_ / 100;
    PixelRGB rgb;

    for(int i = 0; i < h; i++)
        for(int j = 0; j < w; j++)
        {
            rgb = img->at(i, j).RGB();

            // изменение на константу + проверка границ
            rgb.R += l;
            if(rgb.R > 1.0) rgb.R = 1;
            else if(rgb.R < 0.0) rgb.R = 0;

            rgb.G += l;
            if(rgb.G > 1.0) rgb.G = 1;
            else if(rgb.G < 0.0) rgb.G = 0;

            rgb.B += l;
            if(rgb.B > 1.0) rgb.B = 1;
            else if(rgb.B < 0.0) rgb.B = 0;

            img->at(i, j).setRGB(rgb);
        }
    return true;
}
ChangeLightingEffect::ChangeLightingEffect() : lighting_(0)
{
}
ChangeLightingEffect::ChangeLightingEffect(int lighting) : lighting_(lighting)
{
}
int ChangeLightingEffect::lighting()
{
    return lighting_;
}
void ChangeLightingEffect::setLighting(int lighting)
{
    lighting_ = lighting;
}

bool ChangeIntensionEffect::Apply(Image *img)
{
    if(intension_ == 0) return false;
    if(color_ < 0 || color_ > 2) return false;

    int itn = intension_ * 255 / 100;
    int h = img->height();
    int w = img->width();

    for(int i = 0; i < h; i++)
        for(int j = 0; j < w; j++)
        {
            unsigned char* raw = img->atRaw(i, j);

            // изменение + проверка границ
            if((long)raw[color_] + itn > 255) raw[color_] = 255;
            else if((long)raw[color_] + itn < 0) raw[color_] = 0;
            else raw[color_] += itn;
        }
    return true;
}
ChangeIntensionEffect::ChangeIntensionEffect() : intension_(0), color_(0)
{
}
ChangeIntensionEffect::ChangeIntensionEffect(int intension, int color) : intension_(intension), color_(color)
{
}
int ChangeIntensionEffect::intension()
{
    return intension_;
}
int ChangeIntensionEffect::color()
{
    return color_;
}
void ChangeIntensionEffect::setIntension(int intension)
{
    intension_ = intension;
}
void ChangeIntensionEffect::setColor(int color)
{
    color_ = color;
}

bool GrayEffect::Apply(Image *img)
{
    int h = img->height();
    int w = img->width();

    double clr = 0;

    for(int i = 0; i < h; i++)
        for(int j = 0; j < w; j++)
        {
            unsigned char* raw = img->atRaw(i, j);
            clr = (double)raw[0] * 0.3  + (double)raw[1] * 0.59 + (double)raw[2] * 0.11;

            if(clr > 255) clr = 255;
            else if(clr < 0) clr = 0;

            raw[0] = (unsigned char)clr;
            raw[1] = (unsigned char)clr;
            raw[2] = (unsigned char)clr;
        }
    return true;
}
GrayEffect::GrayEffect()
{
}

bool SepiaEffect::Apply(Image *img)
{
    int h = img->height();
    int w = img->width();

    double tr = 0;
    double tg = 0;
    double tb = 0;

    for(int i = 0; i < h; i++)
        for(int j = 0; j < w; j++)
        {
            unsigned char* raw = img->atRaw(i, j);

            tr = 0.393 * (double)raw[0] + 0.769 * (double)raw[1] + 0.189 * (double)raw[2];
            tg = 0.349 * (double)raw[0] + 0.686 * (double)raw[1] + 0.168 * (double)raw[2];
            tb = 0.272 * (double)raw[0] + 0.534 * (double)raw[1] + 0.131 * (double)raw[2];

            if(tr > 255) raw[0] = 255;
            else raw[0] = tr;

            if(tg > 255) raw[1] = 255;
            else  raw[1] = tg;

            if(tb > 255) raw[2] = 255;
            else raw[2] = tb;
        }
    return true;
}
SepiaEffect::SepiaEffect()
{
}