const canvas = document.getElementById('tetris');
const context = canvas.getContext('2d');
context.scale(24, 24); // 1マスを24x24ピクセルで描画

const ROWS = 20;
const COLS = 10;
const EMPTY = 0;
const COLORS = [
  '#222', '#f00', '#0f0', '#00f', '#ff0', '#0ff', '#f0f', '#fa0'
];
// テトリミノの形
const SHAPES = [
  [], // NONE (0)
  [[1,1,1,1]], // I
  [[2,2],[2,2]], // O
  [[0,3,0],[3,3,3]], // T
  [[4,4,0],[0,4,4]], // S
  [[0,5,5],[5,5,0]], // Z
  [[6,0,0],[6,6,6]], // J
  [[0,0,7],[7,7,7]]  // L
];

// フィールド初期化
function createMatrix(rows, cols) {
  const matrix = [];
  while (rows--) matrix.push(new Array(cols).fill(EMPTY));
  return matrix;
}
let field = createMatrix(ROWS, COLS);

// 新しいテトリミノ生成
function randomPiece() {
  const type = (Math.random()*7 | 0) + 1;
  const shape = SHAPES[type];
  return {
    shape: shape.map(row => [...row]),
    x: (COLS/2|0) - (shape[0].length/2|0),
    y: 0,
    type
  };
}
let piece = randomPiece();
let dropCounter = 0;
let dropInterval = 500;
let lastTime = 0;
let gameOver = false;

function collide(field, piece) {
  const {shape, x, y} = piece;
  for (let sy=0; sy<shape.length; sy++) {
    for (let sx=0; sx<shape[sy].length; sx++) {
      if (shape[sy][sx]) {
        let fx = x+sx, fy = y+sy;
        if (fy>=ROWS || fx<0 || fx>=COLS || (fy>=0 && field[fy][fx])) return true;
      }
    }
  }
  return false;
}

function merge(field, piece) {
  piece.shape.forEach((row, sy) => {
    row.forEach((v, sx) => {
      if (v) {
        let fx = piece.x + sx, fy = piece.y + sy;
        if (fy >= 0) field[fy][fx] = piece.type;
      }
    });
  });
}

function rotate(shape) {
  return shape[0].map((_,i) => shape.map(row => row[i]).reverse());
}

function sweep() {
  let lines = 0;
  outer: for (let y=ROWS-1; y>=0; y--) {
    for (let x=0; x<COLS; x++) if (!field[y][x]) continue outer;
    field.splice(y,1);
    field.unshift(new Array(COLS).fill(EMPTY));
    lines++;
    y++;
  }
  return lines;
}

function drawMatrix(matrix, ox, oy) {
  matrix.forEach((row, y) => {
    row.forEach((v, x) => {
      if (v) {
        context.fillStyle = COLORS[v];
        context.fillRect(ox+x, oy+y, 1, 1);
      }
    });
  });
}

function draw() {
  context.fillStyle = COLORS[0];
  context.fillRect(0,0,COLS,ROWS);
  drawMatrix(field, 0, 0);
  drawMatrix(piece.shape, piece.x, piece.y);
}

function update(time = 0) {
  if (gameOver) return;
  const delta = time - lastTime;
  lastTime = time;
  dropCounter += delta;
  if (dropCounter > dropInterval) {
    piece.y++;
    if (collide(field, piece)) {
      piece.y--;
      merge(field, piece);
      sweep();
      piece = randomPiece();
      if (collide(field, piece)) {
        gameOver = true;
        alert('ゲームオーバー!');
        return;
      }
    }
    dropCounter = 0;
  }
  draw();
  requestAnimationFrame(update);
}

document.addEventListener('keydown', e => {
  if (gameOver) return;
  if (e.key==="ArrowLeft") {
    piece.x--;
    if (collide(field, piece)) piece.x++;
  } else if (e.key==="ArrowRight") {
    piece.x++;
    if (collide(field, piece)) piece.x--;
  } else if (e.key==="ArrowDown") {
    piece.y++;
    if (collide(field, piece)) piece.y--;
  } else if (e.key==="ArrowUp") {
    const prev = piece.shape;
    piece.shape = rotate(piece.shape);
    if (collide(field, piece)) piece.shape = prev;
  }
});

update();
