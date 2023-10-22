document.addEventListener('DOMContentLoaded', () => {


  // typing 1
  const typed1 = new Typed(".typing1", {
      strings: [
          "<span>The Best Offline File Converter</span>"
      ],
      typeSpeed: 20,
      showCursor: false,

  });


  const typingElements = document.querySelectorAll(".typing1");
    for (const element of typingElements) {
        element.style.backgroundColor = "rgba(0, 0, 0, 1)";
    }


  // typing 2
  const typed2 = new Typed(".typing2", {
    strings: [
        "<span>.AfterDot</span>"
    ],
    typeSpeed: 60,
    showCursor: false,

});
    




      
  //-----------------------------------------------------------------------frontbg--------------------------------------------------------------------
  
var canvas = document.querySelector('#frontbg'),
ctx = canvas.getContext('2d');

canvas.width = window.innerWidth;
canvas.height = window.innerHeight;

var logoImages = ['pdfimg.png', 'wordimg.png', 'txtimg.png', 'pngimg.png', 'jpgimg.png'];

var imageObjects = [];
var logos = [];

// Load the logo images and create Image objects
function loadImages() {
  for (var i = 0; i < logoImages.length; i++) {
      var img = new Image();
      img.src = logoImages[i];
      imageObjects.push(img);

  }
}

// Function to draw the logos on the canvas
function draw() {
  ctx.fillStyle = 'rgb(60, 2, 135)';
  ctx.fillRect(0, 0, canvas.width, canvas.height);

  for (var i = 0; i < logos.length; i++) {
      var logo = logos[i];
      ctx.drawImage(logo.image, logo.x, logo.y);

      logo.y += logo.speed;

      if (logo.y > canvas.height) {
          logo.y = -logo.image.height;
          logo.x = Math.random() * canvas.width;
      }
  }
}

// Load the images and start the animation

loadImages();


// Create logo instances without repeated images
for (var i = 0; i < logoImages.length; i++) {
  var image = imageObjects[i];
  logos.push({
      image: image,
      x: Math.random() * canvas.width,
      y: -image.height - Math.random() * canvas.height, // Start above the canvas
      speed: 20 // Adjust the speed as needed
  });
}

// Start the animation loop
setInterval(draw, 33);



//-----------------------------------------------------------------------text--------------------------------------------------------------------
  
     
  
  
function loadTextFile(filePath, containerId) {
    fetch(filePath)
      .then(response => {
        if (!response.ok) {
          throw new Error('Network response was not ok');
        }
        return response.text();
      })
      .then(text => {
        const textContainer = document.getElementById(containerId);
  
        // Replace new lines with <br> tags
        text = text.replace(/\n/g, '<br>');
  
        textContainer.innerHTML = text; // Use innerHTML to render <br> tags as line breaks
      })
      .catch(error => {
        console.error('Error:', error);
      });
  }
  
  // Call the function to load and display the text file
  loadTextFile('abttextfile.txt', 'text-container');
  
  
  });
  









