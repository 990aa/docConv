/**
 * Demo JavaScript file to showcase syntax highlighting in DOCX conversion.
 * This file demonstrates various JavaScript/ES6+ features.
 */

// Class with modern syntax
class UserManager {
    constructor(apiUrl) {
        this.apiUrl = apiUrl;
        this.users = [];
    }

    // Async method with arrow function
    async fetchUsers() {
        try {
            const response = await fetch(this.apiUrl);
            const data = await response.json();
            this.users = data.users;
            return this.users;
        } catch (error) {
            console.error('Error fetching users:', error);
            return [];
        }
    }

    // Method with destructuring
    getUserById(id) {
        return this.users.find(user => user.id === id);
    }

    // Arrow function with template literals
    formatUser = (user) => {
        const { name, email, age } = user;
        return `${name} (${email}) - Age: ${age}`;
    }
}

// Higher-order function
const createMultiplier = (factor) => {
    return (number) => number * factor;
};

// Array methods and spread operator
const processData = (numbers) => {
    const doubled = numbers.map(n => n * 2);
    const filtered = doubled.filter(n => n > 10);
    const reduced = filtered.reduce((acc, n) => acc + n, 0);
    
    return {
        doubled: [...doubled],
        filtered,
        sum: reduced
    };
};

// Promise chain
const loadUserData = () => {
    return fetch('https://api.example.com/users')
        .then(response => response.json())
        .then(data => {
            console.log('Users loaded:', data);
            return data;
        })
        .catch(error => {
            console.error('Failed to load users:', error);
            throw error;
        });
};

// Main execution
const main = async () => {
    const manager = new UserManager('https://api.example.com/users');
    
    try {
        const users = await manager.fetchUsers();
        console.log('Total users:', users.length);
        
        // Array operations
        const numbers = [1, 2, 3, 4, 5];
        const result = processData(numbers);
        console.log('Processed data:', result);
        
        // Higher-order function usage
        const triple = createMultiplier(3);
        console.log('Triple of 5:', triple(5));
        
    } catch (error) {
        console.error('Error in main:', error);
    }
};

// Export for module usage
export { UserManager, processData, createMultiplier };
export default main;
