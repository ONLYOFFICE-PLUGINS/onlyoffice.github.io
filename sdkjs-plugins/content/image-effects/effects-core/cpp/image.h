#include <algorithm>

#ifdef __EMSCRIPTEN__
#include <emscripten.h>
#endif // __EMSCRIPTEN__

struct PixelRGB;
struct PixelHSV;
class Pixel;


// основной класс для хранения информации о картинке, хранит "указатель" на нее
// при изменении чего-то тут, меняется и сама картинка
// поэтому нельзя создать "пустой" пиксель, он обязан ссылаться на какую-либо память с картинкой
class Pixel
{
private:

    // 4 bytes
    unsigned char* raw_;

public:
    Pixel() = default;
    Pixel(unsigned char* raw);

    PixelRGB RGB();
    void setRGB(const PixelRGB& rgb);

    unsigned char R() const; // 0
    unsigned char G() const; // 1
    unsigned char B() const; // 2
    unsigned char A() const; // 3

    void setR(unsigned char R);
    void setG(unsigned char G);
    void setB(unsigned char B);
    void setA(unsigned char A);
};

// структура для хранения одного пикселя в формате RGB
// 0 <= R, G, B <= 1
struct PixelRGB
{
    double R;
    double G;
    double B;

    PixelRGB() = default;
    PixelRGB(const Pixel& pixel);

    void fromPixel(const Pixel& pixel);
};

// структура для хранения одного пикселя в формате HSV
// 0 <= H <= 360
// 0 <= S, V <= 1
struct PixelHSV
{
    short H;
    double S;
    double V;

    PixelHSV() = default;
    PixelHSV(const PixelRGB& rgb);

    void fromRGB(const PixelRGB& rgb);
    PixelRGB RGB();
};

// основной класс для хранения картинки
class Image
{
private:

    // указатель на картинку
    unsigned char* raw_;
    int height_;
    int width_;

public:
    Image() = delete;
    Image(const Image& other);
    Image(unsigned char* raw, int height, int width);

    ~Image();

    // вернет пиксель как указатель
    unsigned char* atRaw(int i, int j);

    // вернет пиксель
    Pixel at(int i, int j);

    int height() const;
    int width() const;
};