// Initialize cart
let cart = [];

// Add event listener to form submission
document.getElementById('order-form').addEventListener('submit', async function(event) {
    event.preventDefault();

    const roomNumber = document.getElementById('room-number').value;
    const orderData = {
        roomNumber: roomNumber,
        items: cart.map(item => ({
            name: item.name,
            price: item.price,
            quantity: item.quantity
        }))
    };

    try {
        const response = await fetch('/.netlify/functions/handle-order', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(orderData)
        });

        const result = await response.json();
        alert(result.message); // Notify the user of the response

        // Optionally, clear the cart after successful submission
        cart = [];
        displayCart();
    } catch (error) {
        console.error('Error:', error);
        alert('There was a problem submitting your order.');
    }
});

// Function to add items to the cart
function addToCart(itemName, itemPrice) {
    let item = cart.find(i => i.name === itemName);
    if (item) {
        item.quantity += 1;
    } else {
        cart.push({ name: itemName, price: itemPrice, quantity: 1 });
    }
    displayCart();
}

// Function to display the cart
function displayCart() {
    let cartItemsDiv = document.getElementById('cart-items');
    let cartTotalDiv = document.getElementById('cart-total');
    cartItemsDiv.innerHTML = '';

    let total = 0;
    cart.forEach(item => {
        let itemTotal = item.price * item.quantity;
        total += itemTotal;
        cartItemsDiv.innerHTML += `<p><strong>${item.name}</strong> (x${item.quantity}) - ${itemTotal.toFixed(2)} JD</p>`;
    });

    cartTotalDiv.innerHTML = `Total: ${total.toFixed(2)} JD`;
}
