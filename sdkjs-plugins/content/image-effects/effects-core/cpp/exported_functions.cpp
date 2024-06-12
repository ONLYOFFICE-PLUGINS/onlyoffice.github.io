#include "exported_functions.h"

void change_brightness(unsigned char* data, int h, int w, int brightness)
{
    Image* img = new Image(data, h, w);
    ChangeBrightnessEffect effect(brightness);
    
    effect.Apply(img);
    delete img;
}
void change_saturation(unsigned char* data, int h, int w, int saturation)
{
    Image* img = new Image(data, h, w);
    ChangeSaturationEffect effect(saturation);
    
    effect.Apply(img);
    delete img;
}
void change_transparency(unsigned char* data, int h, int w, int transparency)
{
    Image* img = new Image(data, h, w);
    ChangeTransparencyEffect effect(transparency);
    
    effect.Apply(img);
    delete img;
}
void change_hue(unsigned char* data, int h, int w, int hue)
{
    Image* img = new Image(data, h, w);
    ChangeHueEffect effect(hue);
    
    effect.Apply(img);
    delete img;
}
void change_lighting(unsigned char* data, int h, int w, int lighting)
{
    Image* img = new Image(data, h, w);
    ChangeLightingEffect effect(lighting);
    
    effect.Apply(img);
    delete img;
}

void change_intension_red(unsigned char* data, int h, int w, int intension)
{
    Image* img = new Image(data, h, w);
    ChangeIntensionEffect effect(intension, 0);
    
    effect.Apply(img);
    delete img;
}
void change_intension_green(unsigned char* data, int h, int w, int intension)
{
    Image* img = new Image(data, h, w);
    ChangeIntensionEffect effect(intension, 1);
    
    effect.Apply(img);
    delete img;
}
void change_intension_blue(unsigned char* data, int h, int w, int intension)
{
    Image* img = new Image(data, h, w);
    ChangeIntensionEffect effect(intension, 2);
    
    effect.Apply(img);
    delete img;
}
void shades_of_gray(unsigned char* data, int h, int w, int value)
{
    if(!value) return;

    Image* img = new Image(data, h, w);
    GrayEffect effect;
    
    effect.Apply(img);
    delete img;
}
void sepia(unsigned char* data, int h, int w, int value)
{
    if(!value) return;

    Image* img = new Image(data, h, w);
    SepiaEffect effect;
    
    effect.Apply(img);
    delete img;
}