/* eslint-disable max-len */
/* eslint-disable no-unused-vars */
/* eslint-disable indent */
/**
 * @param {HTMLCanvasElement} mainCanvas
 */
function CBackground(mainCanvas) {
    const canvas = document.createElement('canvas');
    canvas.style.zIndex = '0';
    canvas.style.position = 'absolute';
    canvas.style.display = mainCanvas.style.display;
    mainCanvas.parentElement.appendChild(canvas);
    const context = canvas.getContext('2d');
    resize();
    window.addEventListener('resize', resize);
    if (typeof ResizeObserver !== 'undefined') {
        const ro = new ResizeObserver(function() {
            resize();
        });
        ro.observe(mainCanvas);
    }
    /**
     * TODO: не забыть поменять (тестовая версия)
    */
    function resize() {
        const scale = window.devicePixelRatio;
        const rect = mainCanvas.getBoundingClientRect();
        canvas.style.width = rect.width + 'px';
        canvas.style.height = rect.height + 'px';
        canvas.width = Math.round(scale * rect.right) - Math.round(scale * rect.left);
        canvas.height = Math.round(scale * rect.bottom) - Math.round(scale * rect.top);
        update();
    }
    /**
     */
    function update() {
        const width = (10 * window.devicePixelRatio) >> 0;
        context.fillStyle = '#FFFFFF';
        context.fillRect(0, 0, canvas.width, canvas.height);
        context.beginPath();
        context.fillStyle = '#c9c9c9';
        for (let i = 0; i < canvas.width; i+=width) {
            for (let j = 0; j < canvas.height; j+=width) {
                if ((i / width) % 2 == (j / width) % 2) {
                    context.moveTo(i, j);
                    context.lineTo(i + width, j);
                    context.lineTo(i + width, j + width);
                    context.lineTo(i, j + width);
                    context.lineTo(i, j);
                }
            }
        }
        context.fill();
    }
    this.enable = function() {
        canvas.style.display = mainCanvas.style.display;
        resize();
    };
}
