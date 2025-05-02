// Particle Animation
const canvas = document.getElementById('particleCanvas');
const ctx = canvas.getContext('2d');
canvas.width = window.innerWidth;
canvas.height = window.innerHeight;

const particles = [];
const particleCount = 100;

class Particle {
    constructor() {
        this.x = Math.random() * canvas.width;
        this.y = Math.random() * canvas.height;
        this.size = Math.random() * 3 + 1;
        this.speedX = Math.random() * 0.5 - 0.25;
        this.speedY = Math.random() * 0.5 - 0.25;
        this.opacity = Math.random() * 0.5 + 0.3;
    }

    update() {
        this.x += this.speedX;
        this.y += this.speedY;
        this.opacity = Math.sin(Date.now() * 0.001 + this.x) * 0.3 + 0.5;

        if (this.x < 0 || this.x > canvas.width) this.speedX *= -1;
        if (this.y < 0 || this.y > canvas.height) this.speedY *= -1;
    }

    draw() {
        ctx.beginPath();
        ctx.arc(this.x, this.y, this.size, 0, Math.PI * 2);
        ctx.fillStyle = `rgba(167, 139, 250, ${this.opacity})`;
        ctx.shadowBlur = 10;
        ctx.shadowColor = 'rgba(167, 139, 250, 0.8)';
        ctx.fill();
    }
}

function initParticles() {
    for (let i = 0; i < particleCount; i++) {
        particles.push(new Particle());
    }
}

function animateParticles() {
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    particles.forEach(particle => {
        particle.update();
        particle.draw();
    });
    requestAnimationFrame(animateParticles);
}

initParticles();
animateParticles();

window.addEventListener('resize', () => {
    canvas.width = window.innerWidth;
    canvas.height = window.innerHeight;
});

// Toggle HTML Presentation Type Options
function toggleHtmlOptions() {
    const htmlOptions = document.getElementById('htmlOptions');
    const isHtmlSelected = document.getElementById('html').checked;
    htmlOptions.classList.toggle('hidden', !isHtmlSelected);
    if (isHtmlSelected && !document.querySelector('input[name="html_presentation_type"]:checked')) {
        document.getElementById('minimalist').checked = true;
    }
}

// Form Submission Logic
document.getElementById('generationForm').addEventListener('submit', function (e) {
    e.preventDefault();

    // Show loading indicator
    document.getElementById('loadingIndicator').classList.remove('hidden');
    document.getElementById('resultContainer').classList.add('hidden');
    document.getElementById('errorContainer').classList.add('hidden');
    document.getElementById('generateBtn').disabled = true;
    document.getElementById('generateBtn').classList.add('opacity-70');

    // Get form data
    const formData = new FormData(this);

    // Validate HTML presentation type
    if (formData.get('file_type') === 'html' && !formData.get('html_presentation_type')) {
        formData.set('html_presentation_type', 'minimalist');
    }

    // Send request
    fetch('/generate', {
        method: 'POST',
        body: formData
    })
        .then(response => response.json())
        .then(data => {
            // Hide loading indicator
            document.getElementById('loadingIndicator').classList.add('hidden');
            document.getElementById('generateBtn').disabled = false;
            document.getElementById('generateBtn').classList.remove('opacity-70');

            if (data.success) {
                // Show result
                document.getElementById('resultContainer').classList.remove('hidden');
                document.getElementById('downloadLink').href = data.file_url;

                // Update preview images for HTML presentation types
                if (data.preview_images) {
                    const minimalistImg = document.querySelector('#htmlOptions img[alt="Minimalist Preview"]');
                    const modernImg = document.querySelector('#htmlOptions img[alt="Modern Preview"]');
                    const professionalImg = document.querySelector('#htmlOptions img[alt="Professional Preview"]');

                    if (data.preview_images.minimalist) minimalistImg.src = data.preview_images.minimalist;
                    if (data.preview_images.modern) modernImg.src = data.preview_images.modern;
                    if (data.preview_images.professional) professionalImg.src = data.preview_images.professional;
                }

                // Show preview
                const previewContent = document.getElementById('previewContent');
                previewContent.innerHTML = '';

                if (data.preview && data.preview.presentation && data.preview.presentation.slides) {
                    const slides = data.preview.presentation.slides;
                    slides.forEach((slide, index) => {
                        const slideElement = document.createElement('div');
                        if (index > 0) {
                            slideElement.classList.add('mt-4', 'pt-4', 'border-t', 'border-gray-200');
                        }

                        const titleElement = document.createElement('h4');
                        titleElement.textContent = slide.title || 'Untitled Slide';
                        titleElement.classList.add('font-semibold', 'text-gray-900', 'mb-3');
                        slideElement.appendChild(titleElement);

                        if (slide.bullets && slide.bullets.length) {
                            const bulletList = document.createElement('ul');
                            bulletList.classList.add('list-disc', 'pl-6', 'space-y-2');

                            slide.bullets.forEach(bullet => {
                                const bulletItem = document.createElement('li');
                                bulletItem.textContent = bullet;
                                bulletItem.classList.add('text-gray-700', 'text-sm');
                                bulletList.appendChild(bulletItem);
                            });

                            slideElement.appendChild(bulletList);
                        }

                        previewContent.appendChild(slideElement);

                        // Add note about image slide in PowerPoint only
                        if (formData.get('file_type') === 'pptx') {
                            const imageNote = document.createElement('p');
                            imageNote.textContent = "Will include Pexels image slide after this content";
                            imageNote.classList.add('text-xs', 'text-indigo-600', 'mt-2', 'italic');
                            slideElement.appendChild(imageNote);
                        }

                        if (formData.get('file_type') === 'html') {
                            const htmlNote = document.createElement('p');
                            htmlNote.textContent = `HTML version includes interactive navigation (${formData.get('html_presentation_type')} style)`;
                            htmlNote.classList.add('text-xs', 'text-indigo-600', 'mt-2', 'italic');
                            slideElement.appendChild(htmlNote);
                        }
                    });
                } else {
                    previewContent.textContent = "No preview available";
                }
            } else {
                // Show error
                document.getElementById('errorContainer').classList.remove('hidden');
                document.getElementById('errorMessage').textContent = data.error || "Something went wrong";
            }
        })
        .catch(error => {
            // Hide loading indicator and show error
            document.getElementById('loadingIndicator').classList.add('hidden');
            document.getElementById('errorContainer').classList.remove('hidden');
            document.getElementById('errorMessage').textContent = "Network error occurred";
            document.getElementById('generateBtn').disabled = false;
            document.getElementById('generateBtn').classList.remove('opacity-70');
            console.error('Error:', error);
        });
});

function updateTooltip(value) {
    const tooltip = document.getElementById("tooltip");
    tooltip.textContent = value || "Your input will show here...";
}