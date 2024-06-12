#include <iostream>
#include "effect.h"

#ifdef __cplusplus
extern "C"
{
#endif // __cplusplus

void change_brightness(unsigned char* data, int h, int w, int brightness);
void change_saturation(unsigned char* data, int h, int w, int saturation);
void change_hue(unsigned char* data, int h, int w, int hue);

void change_transparency(unsigned char* data, int h, int w, int transparency);
void change_lighting(unsigned char* data, int h, int w, int lighting);

void change_intension_red(unsigned char* data, int h, int w, int intension);
void change_intension_green(unsigned char* data, int h, int w, int intension);
void change_intension_blue(unsigned char* data, int h, int w, int intension);

void shades_of_gray(unsigned char* data, int h, int w, int value);
void sepia(unsigned char* data, int h, int w, int value);

#ifdef __cplusplus
}
#endif // __cplusplus

