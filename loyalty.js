let points = 0;

function addPoints(amount) {
  points += 100;
  document.getElementById("points").innerHTML = `Points: ${points}`;
}
function redeemPoints() {
  if (points >= 100) {
    points -= 100;
    document.getElementById("points").innerHTML = `Points: ${points}`;
    alert("You have redeemed 100 points!");
  } else {
    alert("You don't have enough points to redeem!");
  }
}

function addToWallet() {
  let wallet = document.getElementById("wallet").value;
  points += parseInt(wallet);
  document.getElementById("points").innerHTML = `Points: ${points}`;
}

function registerPurchase() {
  let amount = document.getElementById("amount").value;
  addPoints(amount);
  alert("Purchase registered!");
}