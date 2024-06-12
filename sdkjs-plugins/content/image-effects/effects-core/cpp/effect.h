#include "image.h"

// интерфейс для эффекта
class IEffect
{
public:
    virtual bool Apply(Image* img) = 0;
};

// добавляет яркость (V в HSV)
class ChangeBrightnessEffect : public IEffect
{
private:
    int brightness_;
    
public:
    virtual bool Apply(Image *img);

    ChangeBrightnessEffect();
    ChangeBrightnessEffect(int brightness);

    int brightness();
    void setBrightness(int brightness);
};

// мзменяет тон (H в HSV)
class ChangeHueEffect : public IEffect
{
private:
    int hue_;
    
public:
    virtual bool Apply(Image *img);

    ChangeHueEffect();
    ChangeHueEffect(int hue);

    int hue();
    void setHue(int hue);
};

// добавляет насыщенность (S в HSV)
class ChangeSaturationEffect : public IEffect
{
private:
    int saturation_;
    
public:
    virtual bool Apply(Image *img);

    ChangeSaturationEffect();
    ChangeSaturationEffect(int saturation);

    int saturation();
    void setSaturation(int saturation);
};

// меняет прозрачность (A в RGBA)
class ChangeTransparencyEffect : public IEffect
{
private:
    int transparency_;

public:
    virtual bool Apply(Image *img);

    ChangeTransparencyEffect();
    ChangeTransparencyEffect(int transparency);

    int transparency();
    void setTransparency(int transparency);
};

// осветление картинки (изменяем rgb на константу)
class ChangeLightingEffect : public IEffect
{
private:
    int lighting_;

public:
    virtual bool Apply(Image *img);

    ChangeLightingEffect();
    ChangeLightingEffect(int lighting);

    int lighting();
    void setLighting(int lighting);
};

// меняет "количество" цвета (R или G или B на константу)
class ChangeIntensionEffect : public IEffect
{
private:
    int intension_;
    int color_;

public:
    virtual bool Apply(Image *img);

    ChangeIntensionEffect();

    // 0 - R, 1 - G, 2 - B
    ChangeIntensionEffect(int intension, int color);

    int intension();
    int color();

    void setIntension(int intension);
    void setColor(int color);
};

// оттенки серого
class GrayEffect : public IEffect
{
private:

public:
    virtual bool Apply(Image *img);
    GrayEffect();
};

// сепия
class SepiaEffect : public IEffect
{
private:

public:
    virtual bool Apply(Image *img);
    SepiaEffect();
};