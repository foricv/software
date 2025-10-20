document.addEventListener("DOMContentLoaded", () => {
  const imageInput = document.getElementById("imageInput");
  const prepareBtn = document.getElementById("prepareBtn");

  let selectedFiles = [];

  // Track rotation per image
  const rotations = { img1: 0, img2: 0, img3: 0 };

  // Track fit/fill mode per image (default: fit)
  const fitFillModes = { img1: 'fit', img2: 'fit', img3: 'fit' };

  imageInput.addEventListener("change", (e) => {
    selectedFiles = Array.from(e.target.files).filter(f =>
      /\.(jpe?g|png)$/i.test(f.name)
    );
  });

  prepareBtn.addEventListener("click", () => {
    if (selectedFiles.length !== 3) {
      alert("Please select exactly 3 image files: CNIC Front, CNIC Back, Full Photo.");
      return;
    }

    prepareBtn.disabled = true;

    selectedFiles.forEach((file, index) => {
      if (index > 2) return;
      const reader = new FileReader();
      reader.onload = (e) => {
        const img = document.getElementById(`img${index + 1}`);
        img.src = e.target.result;
        img.style.filter = "";
        rotations[`img${index + 1}`] = 0;
        img.style.transform = "rotate(0deg)";
        // Reset fit/fill to fit (contain)
        fitFillModes[`img${index + 1}`] = 'fit';
        img.style.objectFit = 'contain';
        // Reset grayscale if any
        img.style.filter = '';
        // Reset button text for fit/fill buttons
        const btn = img.parentElement.querySelector("button[onclick^='toggleFitFill']");
        if (btn) btn.textContent = 'Fill Container';
      };
      reader.readAsDataURL(file);
    });

    // Enable button after loading images (small delay to let images load)
    setTimeout(() => {
      prepareBtn.disabled = false;
    }, 500);
  });

  // Cropper logic
  let currentTargetImg = null;
  let cropper = null;

  window.openCropModal = function (imgId) {
    currentTargetImg = document.getElementById(imgId);
    const cropImage = document.getElementById("cropImage");
    cropImage.src = currentTargetImg.src;
    document.getElementById("cropModal").style.display = "flex";
    if (cropper) cropper.destroy();
    cropper = new Cropper(cropImage, { viewMode: 1 });
  };

  document.getElementById("applyCrop").onclick = () => {
    if (cropper && currentTargetImg) {
      const canvas = cropper.getCroppedCanvas();
      currentTargetImg.src = canvas.toDataURL();
    }
    cropper?.destroy();
    document.getElementById("cropModal").style.display = "none";
  };

  document.getElementById("cancelCrop").onclick = () => {
    cropper?.destroy();
    document.getElementById("cropModal").style.display = "none";
  };

  // Toggle grayscale on/off
  window.applyGrayscale = function (imgId) {
    const img = document.getElementById(imgId);
    const currentFilter = img.style.filter || '';
    if (currentFilter.includes('grayscale(100%)')) {
      img.style.filter = currentFilter.replace(/grayscale\(100%\)/g, '').trim();
    } else {
      img.style.filter = (currentFilter + ' grayscale(100%)').trim();
    }
  };

  // Adjust Modal Logic
  window.openAdjustModal = function (imgId) {
    currentTargetImg = document.getElementById(imgId);
    const modal = document.getElementById("adjustModal");
    const adjustImage = document.getElementById("adjustImage");

    adjustImage.src = currentTargetImg.src;
    document.getElementById("brightness").value = 100;
    document.getElementById("contrast").value = 100;
    adjustImage.style.filter = "brightness(100%) contrast(100%)";

    modal.style.display = "flex";
  };

  document.getElementById("brightness").oninput =
  document.getElementById("contrast").oninput = () => {
    const b = document.getElementById("brightness").value;
    const c = document.getElementById("contrast").value;
    document.getElementById("adjustImage").style.filter = `brightness(${b}%) contrast(${c}%)`;
  };

  document.getElementById("applyAdjust").onclick = () => {
    const canvas = document.createElement("canvas");
    const ctx = canvas.getContext("2d");
    const img = document.getElementById("adjustImage");

    const tempImg = new Image();
    tempImg.crossOrigin = "anonymous";
    tempImg.onload = () => {
      canvas.width = tempImg.width;
      canvas.height = tempImg.height;

      const brightness = document.getElementById("brightness").value;
      const contrast = document.getElementById("contrast").value;

      ctx.filter = `brightness(${brightness}%) contrast(${contrast}%)`;
      ctx.drawImage(tempImg, 0, 0);
      currentTargetImg.src = canvas.toDataURL();

      document.getElementById("adjustModal").style.display = "none";
    };
    tempImg.src = img.src;
  };

  document.getElementById("cancelAdjust").onclick = () => {
    document.getElementById("adjustModal").style.display = "none";
    document.getElementById("brightness").value = 100;
    document.getElementById("contrast").value = 100;
  };

  // Rotate function
  window.rotateImage = function (imgId) {
    rotations[imgId] = (rotations[imgId] + 90) % 360;
    const img = document.getElementById(imgId);
    img.style.transform = `rotate(${rotations[imgId]}deg)`;
  };

  // Fit/Fill toggle function
window.toggleFitFill = function (imgId, btn) {
  const img = document.getElementById(imgId);
  const buttons = img.parentElement.querySelectorAll('.buttons button'); // Get all buttons within the buttons container

  if (!img) return;

  if (fitFillModes[imgId] === 'fit') {
    // Switch to fill (stretch)
    img.style.objectFit = 'fill';
    fitFillModes[imgId] = 'fill';
    // Hide all buttons except for the one that triggered the function
    buttons.forEach(button => {
      if (button !== btn) {
        button.style.display = 'none';
      }
    });
    btn.textContent = 'Fit Aspect Ratio';  // Change button text to 'Fit Aspect Ratio'
  } else {
    // Switch to fit (contain)
    img.style.objectFit = 'contain';
    fitFillModes[imgId] = 'fit';
    // Show all buttons again when in 'fit' mode
    buttons.forEach(button => {
      button.style.display = '';  // Reset display to default
    });
    btn.textContent = 'Fill Container';  // Change button text back to 'Fill Container'
  }
};

  // Delete Image
  window.deleteImage = function (imgId) {
    const img = document.getElementById(imgId);
    img.src = "";  // Clear the image source
    const buttons = img.parentElement.querySelector('.buttons');
    buttons.style.display = 'block';  // Ensure buttons are visible again if image is deleted
  };

});

 // Toggle between templates
  const radioButtons = document.querySelectorAll('input[name="layoutMode"]');
  const template3 = document.getElementById('template3');
  const template8 = document.getElementById('template8');

  radioButtons.forEach(radio => {
    radio.addEventListener('change', () => {
      if (radio.value === 'template3') {
        template3.style.display = 'block';
        template8.style.display = 'none';
      } else if (radio.value === 'template8') {
        template3.style.display = 'none';
        template8.style.display = 'block';
      }
    });
  });

document.addEventListener("DOMContentLoaded", () => {
  const multiImageInput = document.getElementById("multiImageInput");
  const generateTemplateBtn = document.getElementById("generateTemplateBtn");
  const multiLayout = document.getElementById("multiLayout");

  generateTemplateBtn.addEventListener("click", () => {
    const files = Array.from(multiImageInput.files).filter(f =>
      /\.(jpe?g|png)$/i.test(f.name)
    );

    const copies = parseInt(document.getElementById("copiesPerImage").value) || 1;

    if (files.length === 0 || copies < 1) {
      alert("Please select images and set number of copies.");
      return;
    }

    if (files.length % 2 !== 0) {
      alert("Please upload CNIC Front/Back images in pairs.");
      return;
    }

    const expectedBlocks = files.length * copies;
    if (expectedBlocks !== 8) {
      if (!confirm(`You are generating ${expectedBlocks} blocks. Proceed?`)) return;
    }

    const fileReaders = files.map(file => {
      return new Promise(resolve => {
        const reader = new FileReader();
        reader.onload = e => resolve(e.target.result);
        reader.readAsDataURL(file);
      });
    });

    Promise.all(fileReaders).then(results => {
      const imagePairs = [];
      for (let i = 0; i < results.length; i += 2) {
        imagePairs.push([results[i], results[i + 1]]); // front/back
      }

      const duplicatedPairs = [];
      imagePairs.forEach(pair => {
        for (let i = 0; i < copies; i++) {
          duplicatedPairs.push(pair);
        }
      });

      renderMultiBlocks(duplicatedPairs.slice(0, 4)); // only 4 rows (8 blocks max)
    });
  });

  function renderMultiBlocks(pairs) {
    multiLayout.innerHTML = '';

    pairs.forEach(([front, back]) => {
      const row = document.createElement('div');
      row.style.display = 'flex';
      row.style.gap = '20px';
      row.style.marginBottom = '20px';

      [front, back].forEach(src => {
        const container = document.createElement('div');
        container.className = 'cnic-block';
        const img = document.createElement('img');
        img.src = src;
        img.style.objectFit = 'contain';
        container.appendChild(img);
        row.appendChild(container);
      });

      multiLayout.appendChild(row);
    });
  }
});
