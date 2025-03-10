
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>🚀 Excel Mastery: Basic to Advanced Course 📊</title>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/tailwindcss/2.2.19/tailwind.min.css">
    <script src="https://cdnjs.cloudflare.com/ajax/libs/alpinejs/3.12.0/cdn.min.js" defer></script>
</head>
<body class="bg-black text-white">
    <header class="bg-green-600 text-white py-6 text-center">
        <h1 class="text-4xl font-bold">🚀 Excel Mastery: Basic to Advanced Course 📊</h1>
        <p class="mt-2 text-lg">Master Excel from scratch and boost your career! 💼📈</p>
        <p class="mt-2 text-lg">Master Excel from scratch and boost your career! 🌟📈</p>
    </header>
    
    <section class="container mx-auto p-6 text-center" x-data="{
        reasons: [
            { title: '🔥 Beginner Friendly', description: 'Start with the basics and build a solid foundation. 🏗️' },
            { title: '🚀 Advanced Techniques', description: 'Master formulas, pivot tables, and automation. 🛠️' },
            { title: '🎓 Certificate of Completion', description: 'Get certified and boost your resume. ✅' }
        ]
    }">
        <h2 class="text-3xl font-semibold">🌟 Why Join This Course? 🌟</h2>
        <p class="mt-4 text-lg">Learn formulas, pivot tables, automation with VBA, and data analysis skills. 📊</p>
        <div class="mt-6 grid grid-cols-1 md:grid-cols-3 gap-6">
            <template x-for="(reason, index) in reasons" :key="index">
                <div class="p-4 bg-gray-800 shadow-lg rounded-lg" x-transition>
                    <h3 class="text-xl font-bold" x-text="reason.title"></h3>
                    <p x-text="reason.description"></p>
                </div>
            </template>
        </div>
    </section>
    
    <section class="bg-gray-900 p-6 text-center">
        <h2 class="text-3xl font-semibold">📚 Course Curriculum 📚</h2>
        <ul class="mt-4 text-lg text-left mx-auto w-3/4" x-data="{topics: ['✨ Introduction to Excel', '📊 Formulas & Functions', '📈 Data Visualization & Pivot Tables', '⚡ Advanced Excel Functions', '🤖 Macros & VBA Automation']}">
            <template x-for="topic in topics" :key="topic">
                <li x-transition>✔️ <span x-text="topic"></span></li>
            </template>
        </ul>
    </section>
    
    <section class="bg-gray-800 p-6 text-center" x-data="{testimonials: [{name: 'Rahul M.', text: '🔥 This course changed my career! Highly recommended.'}, {name: 'Priya S.', text: '💡 Excellent teaching, practical examples, and great support.'}]}">
        <h2 class="text-3xl font-semibold">💬 What Our Students Say 💬</h2>
        <div class="mt-6 grid grid-cols-1 md:grid-cols-2 gap-6">
            <template x-for="testimonial in testimonials" :key="testimonial.name">
                <div class="p-4 bg-gray-700 shadow-lg rounded-lg" x-transition>
                    <p x-text="testimonial.text"></p>
                    <p class="font-bold" x-text="'- ' + testimonial.name"></p>
                </div>
            </template>
        </div>
    </section>
    
    <section class="bg-green-600 text-white p-6 text-center">
        <h2 class="text-3xl font-semibold">🚀 Enroll Now 🎯</h2>
        <p class="mt-2 text-lg">Start your journey to Excel mastery today. 🎓📊</p>
        
        <form class="mt-4" action="#" method="POST">
            <input type="text" name="name" placeholder="👤 Your Name" class="p-3 rounded-lg text-black w-64 mb-3" required>
            <input type="email" name="email" placeholder="📧 Your Email" class="p-3 rounded-lg text-black w-64 mb-3" required>
            <input type="tel" name="mobile" placeholder="📱 Your Mobile Number" class="p-3 rounded-lg text-black w-64 mb-3" required>
            <button type="submit" class="bg-white text-green-600 px-6 py-3 rounded-lg font-bold text-xl transition-transform transform hover:scale-105">🎟️ Sign Up</button>
        </form>
    </section>

    <section class="bg-gray-900 text-white p-6 text-center">
        <h2 class="text-3xl font-semibold">📲 Connect With Us 📲</h2>
        <p class="mt-2 text-lg">Follow us on social media for updates and tips! 🌐</p>
        <div class="mt-4 flex justify-center gap-4">
            <a href="#" class="bg-blue-600 px-6 py-3 rounded-lg text-white font-bold text-xl transition-transform transform hover:scale-105">🔵 Facebook</a>
            <a href="#" class="bg-blue-400 px-6 py-3 rounded-lg text-white font-bold text-xl transition-transform transform hover:scale-105">🐦 Twitter</a>
            <a href="#" class="bg-pink-500 px-6 py-3 rounded-lg text-white font-bold text-xl transition-transform transform hover:scale-105">📸 Instagram</a>
            <a href="#" class="bg-red-600 px-6 py-3 rounded-lg text-white font-bold text-xl transition-transform transform hover:scale-105">🎥 YouTube</a>
            <a href="#" class="bg-green-500 px-6 py-3 rounded-lg text-white font-bold text-xl transition-transform transform hover:scale-105">👥 WhatsApp</a>
        </div>
    </section>

    <section class="bg-gray-800 text-white p-6 text-center">
        <h2 class="text-3xl font-semibold">💳 Secure Payment Options 💳</h2>
        <p class="mt-2 text-lg">Choose your preferred payment method and enroll today! ✅</p>
        <div class="mt-4 flex justify-center gap-4">
            <a href="#" class="bg-blue-700 px-6 py-3 rounded-lg text-white font-bold text-xl transition-transform transform hover:scale-105">💳 Credit Card</a>
            <a href="#" class="bg-blue-500 px-6 py-3 rounded-lg text-white font-bold text-xl transition-transform transform hover:scale-105">💳 Debit Card</a>
            <a href="#" class="bg-green-600 px-6 py-3 rounded-lg text-white font-bold text-xl transition-transform transform hover:scale-105">📲 UPI</a>
            <a href="#" class="bg-indigo-600 px-6 py-3 rounded-lg text-white font-bold text-xl transition-transform transform hover:scale-105">💰 PayPal</a>
        </div>
    </section>
    
    <footer class="bg-gray-900 text-white p-4 text-center">
        <p>© 2025 Excel Mastery. All rights reserved. 🚀📊</p>
    </footer>
</body>
</html>
