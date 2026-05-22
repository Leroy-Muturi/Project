const dragon = document.getElementById('dragon');
const fire = document.getElementById('fire');

const state = {
  mouseX: window.innerWidth / 2,
  mouseY: window.innerHeight / 2,
  dragonX: window.innerWidth / 2 - 120,
  dragonY: window.innerHeight / 2 - 80,
};

const chaseOffset = 85;
const speed = 0.08;

window.addEventListener('mousemove', (event) => {
  state.mouseX = event.clientX;
  state.mouseY = event.clientY;
});

window.addEventListener('click', () => {
  fire.classList.remove('active');
  void fire.offsetWidth;
  fire.classList.add('active');
});

function animate() {
  const dx = state.mouseX - state.dragonX;
  const dy = state.mouseY - state.dragonY;
  const angle = Math.atan2(dy, dx);

  const targetX = state.mouseX - Math.cos(angle) * chaseOffset;
  const targetY = state.mouseY - Math.sin(angle) * chaseOffset;

  state.dragonX += (targetX - state.dragonX) * speed;
  state.dragonY += (targetY - state.dragonY) * speed;

  dragon.style.transform = `translate(${state.dragonX}px, ${state.dragonY}px) rotate(${angle}rad)`;

  requestAnimationFrame(animate);
}

animate();
