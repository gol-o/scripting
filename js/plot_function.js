var rgb = (r, g, b) => {
    return (r & 0xF0 ? '#' : '#0') + (r << 16 | g << 8 | b).toString(16)
}

var angle = 0;

setInterval(() => {
    document.body.style.color = rgb(angle * 10 % 255, 0, 0);

    if (angle < 20)
        document.write("■".repeat(Math.abs(50 * Math.sqrt(angle * 6 * Math.PI / 180))) + "<p>");
    else if (angle < 50)
        document.write("■".repeat(Math.abs(50 * Math.sin(angle * 6 * Math.PI / 180))) + "<p>");
    else if (angle < 80)
        document.write("■".repeat(Math.abs(50 * Math.cos(angle * 6 * Math.PI / 180))) + "<p>");
    else if (angle < 120)
        document.write("■".repeat(Math.abs(50 * Math.tan(angle * 6 * Math.PI / 180) % 50)) + "<p>");
    else
        document.write("■".repeat(Math.abs(50 * 1 / Math.tan(angle * 6 * Math.PI / 180) % 50)) + "<p>");
    
    window.scrollTo(0, angle * 100);

    angle++;
    angle = angle % 360;
}, 50);

