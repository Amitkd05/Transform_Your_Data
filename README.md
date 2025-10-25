<head>
    <meta charset="UTF-8">
    <meta http-equiv="Content-Security-Policy" content="default-src 'self'; script-src 'self' 'unsafe-inline' https://cdn.tailwindcss.com https://cdn.jsdelivr.net/npm/; style-src 'self' 'unsafe-inline' https://fonts.googleapis.com; font-src 'self' https://fonts.gstatic.com; img-src 'self' https://placehold.co; connect-src https://api.emailjs.com;">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Excel Speed & Formula Vault</title>

    <!-- SEO & Social Sharing Meta Tags -->
    <meta name="description" content="Your ultimate resource for mastering Excel. A searchable, interactive vault of hundreds of Excel formulas and shortcuts, complete with easy examples.">
    <link rel="canonical" href="https://your-website-url.com/" />

    <!-- Open Graph / Facebook -->
    <meta property="og:type" content="website">
    <meta property="og:url" content="https://your-website-url.com/">
    <meta property="og:title" content="Excel Speed & Formula Vault">
    <meta property="og:description" content="The ultimate resource for mastering Excel. Search hundreds of formulas and shortcuts with easy examples.">
    <meta property="og:image" content="https://placehold.co/1200x630/1e293b/ffffff?text=Excel%20Speed%20%26%20Formula%20Vault">
    <meta property="og:site_name" content="Excel Speed & Formula Vault">

    <!-- Twitter -->
    <meta property="twitter:card" content="summary_large_image">
    <meta property="twitter:url" content="https://your-website-url.com/">
    <meta property="twitter:title" content="Excel Speed & Formula Vault">
    <meta property="twitter:description" content="The ultimate resource for mastering Excel. Search hundreds of formulas and shortcuts with easy examples.">
    <meta property="twitter:image" content="https://placehold.co/1200x630/1e293b/ffffff?text=Excel%20Speed%20%26%20Formula%20Vault">

    <script src="https://cdn.tailwindcss.com"></script>
    <script>
        tailwind.config = { darkMode: 'class' }
    </script>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
    <!-- EmailJS Library -->
    <script type="text/javascript" src="https://cdn.jsdelivr.net/npm/@emailjs/browser@3/dist/email.min.js"></script>
    <style>
        body {
            font-family: 'Inter', sans-serif;
        }
        /* Custom scrollbar for better aesthetics */
        ::-webkit-scrollbar { width: 8px; }
        ::-webkit-scrollbar-track { background-color: #1f2937; }
        ::-webkit-scrollbar-thumb { background-color: #4b5563; border-radius: 10px; }

        .copied-toast {
            position: fixed; bottom: 20px; left: 50%;
            transform: translateX(-50%); padding: 10px 20px;
            border-radius: 8px; color: white; background-color: #22c55e;
            opacity: 0; transition: opacity 0.3s ease, transform 0.3s ease;
            z-index: 50; pointer-events: none;
        }
        .copied-toast.show {
            opacity: 1; transform: translateX(-50%) translateY(-10px);
        }
        .favorite-btn:hover svg {
            transform: scale(1.2);
        }
        .favorite-btn.favorited svg {
            fill: #facc15; /* yellow-400 */
            stroke: #facc15;
        }

        /* Modal styles */
        .modal-base {
            transition: opacity 0.2s ease-in-out;
        }
        .modal-content {
            transform: scale(0.95);
            transition: transform 0.2s ease-in-out;
        }
        .modal-base.flex {
            opacity: 1;
        }
        .modal-base.flex .modal-content {
            transform: scale(1);
        }

        /* --- NEW 3D Floating Card Style --- */
        .formula-card, .testimonial-card {
            will-change: transform, box-shadow; /* Performance optimization */
            transition: transform 0.25s ease-out, box-shadow 0.25s ease-out;
            border: 1px solid #e5e7eb; /* Added border for light mode */
        }
        .dark .formula-card, .dark .testimonial-card {
            border-color: #374151; /* Adjusted border for dark mode */
        }
        .formula-card:hover, .testimonial-card:hover {
            transform: translateY(-12px) scale(1.03); /* More lift and zoom */
            /* Deeper shadow and a more prominent, wider glow */
            box-shadow: 0 30px 50px -15px rgb(0 0 0 / 0.25), 0 0 35px 0px rgb(59 130 246 / 0.7);
        }
        .dark .formula-card:hover, .dark .testimonial-card:hover {
             /* Adjusted prominent glow for dark mode */
             box-shadow: 0 30px 50px -15px rgb(0 0 0 / 0.50), 0 0 35px 0px rgb(96 165 250 / 0.7);
        }

        /* --- NEW 3D Category Button Style --- */
        .category-btn {
            will-change: transform, box-shadow;
            transition: transform 0.2s ease-out, box-shadow 0.2s ease-out;
        }
        
        .category-btn:not(.active-category):hover {
            transform: translateY(-4px) scale(1.05);
            box-shadow: 0 10px 20px -8px rgb(0 0 0 / 0.20), 0 0 15px -5px rgb(59 130 246 / 0.5);
        }

        .dark .category-btn:not(.active-category):hover {
            box-shadow: 0 10px 20px -8px rgb(0 0 0 / 0.40), 0 0 15px -5px rgb(96 165 250 / 0.5);
        }

        /* --- NEW Alphabet Filter Style --- */
        .alphabet-btn {
            transition: background-color 0.2s ease-out, color 0.2s ease-out, transform 0.2s ease-out, box-shadow 0.2s ease-out;
            will-change: transform, box-shadow;
        }

        .alphabet-btn:not(.active-alphabet):hover {
            transform: translateY(-4px) scale(1.05);
            box-shadow: 0 10px 15px -5px rgb(0 0 0 / 0.15), 0 0 10px -5px rgb(59 130 246 / 0.5);
        }
        
        .dark .alphabet-btn:not(.active-alphabet):hover {
             box-shadow: 0 10px 15px -5px rgb(0 0 0 / 0.3), 0 0 10px -5px rgb(96 165 250 / 0.5);
        }
        
        /* --- NEW 3D Floating Tab Bar --- */
        .tab-btn.active-tab {
            color: #2563eb; /* blue-600 */
        }
        .dark .tab-btn.active-tab {
            color: #60a5fa; /* blue-400 */
        }

        /* Policy modal content scrollbar */
        .policy-content::-webkit-scrollbar { width: 5px; }
        .policy-content::-webkit-scrollbar-track { background-color: #e5e7eb; border-radius: 10px; }
        .dark .policy-content::-webkit-scrollbar-track { background-color: #374151; }
        .policy-content::-webkit-scrollbar-thumb { background-color: #9ca3af; border-radius: 10px; }
        .dark .policy-content::-webkit-scrollbar-thumb { background-color: #6b7280; }

    </style>
<style>
  /* Enables natural momentum scrolling on iOS */
  #categoryDropdownList {
    -webkit-overflow-scrolling: touch;
  }
</style>

</head>
<body class="bg-gray-100 dark:bg-gray-900 text-gray-800 dark:text-gray-200 transition-colors duration-300">

    <!-- Main Container -->
    <div class="container mx-auto p-4 sm:p-6 lg:p-8">

        <!-- Header -->
        <header class="mb-8 flex flex-col sm:flex-row sm:justify-between sm:items-start sm:gap-4">
            <div class="text-center sm:text-left order-2 sm:order-1 mt-4 sm:mt-0">
                <h1 class="text-4xl sm:text-5xl font-bold text-gray-900 dark:text-white">Excel Speed & Formula Vault</h1>
                <p class="text-lg text-gray-600 dark:text-gray-400 mt-2">Your ultimate resource for mastering Excel</p>
                <h2 id="greeting" class="text-2xl sm:text-3xl font-semibold text-blue-600 dark:text-blue-400 mt-4"></h2>
            </div>
            <div id="datetime-container" class="w-full max-w-xs sm:w-auto order-1 sm:order-2 self-center sm:self-start flex-shrink-0 p-2 sm:p-3 bg-white/30 dark:bg-black/30 backdrop-blur-sm rounded-xl shadow-lg text-right border border-white/20">
                <div id="time-display" class="text-xl sm:text-2xl font-bold text-gray-800 dark:text-white"></div>
                <div id="date-display" class="text-xs sm:text-sm text-gray-600 dark:text-gray-400"></div>
            </div>
        </header>

<!-- Tabs -->
        <div id="tabs-container" class="relative bg-gray-200 dark:bg-gray-800 p-1 rounded-full shadow-inner mb-8 w-full max-w-md mx-auto flex items-center justify-around">
            <div id="active-tab-slider" class="absolute top-1 bottom-1 left-0 bg-white dark:bg-gray-700 rounded-full shadow-md transition-all duration-300 ease-in-out"></div>
            <button data-tab="formulas" class="tab-btn active-tab relative z-10 w-full px-4 py-2 text-base sm:text-lg font-medium">Formulas</button>
            <button data-tab="shortcuts" class="tab-btn relative z-10 w-full px-4 py-2 text-base sm:text-lg font-medium">Shortcuts</button>
            <button data-tab="favorites" class="tab-btn relative z-10 w-full px-4 py-2 text-base sm:text-lg font-medium flex items-center justify-center gap-1">
                <svg xmlns="http://www.w3.org/2000/svg" class="w-5 h-5" viewBox="0 0 20 20" fill="currentColor"><path d="M9.049 2.927c.3-.921 1.603-.921 1.902 0l1.07 3.292a1 1 0 00.95.69h3.462c.969 0 1.371 1.24.588 1.81l-2.8 2.034a1 1 0 00-.364 1.118l1.07 3.292c.3.921-.755 1.688-1.54 1.118l-2.8-2.034a1 1 0 00-1.175 0l-2.8 2.034c-.784.57-1.838-.197-1.539-1.118l1.07-3.292a1 1 0 00-.364-1.118L2.98 8.72c-.783-.57-.38-1.81.588-1.81h3.461a1 1 0 00.951-.69l1.07-3.292z" /></svg>
                Favorites
            </button>
        </div>

        <!-- Controls: Search, Tabs, and Theme Toggle -->
        <div class="flex flex-col sm:flex-row justify-between items-center mb-6 gap-4">
            <div class="relative w-full sm:w-auto flex-grow">
                <input type="text" id="searchInput" placeholder="Search hundreds of formulas & shortcuts..." class="w-full pl-10 pr-4 py-2 border rounded-full bg-white dark:bg-gray-800 border-gray-300 dark:border-gray-600 focus:outline-none focus:ring-2 focus:ring-blue-500">
                <svg class="absolute left-3 top-1/2 -translate-y-1/2 w-5 h-5 text-gray-400" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z"></path></svg>
            </div>

	    <!-- NEW: All-Categories Dropdown -->
<div class="relative w-full sm:w-72 mt-2 sm:mt-0">
  <button id="categoryDropdownBtn"
	  aria-haspopup="listbox"
          aria-expanded="false"
          aria-controls="categoryDropdownList"
          class="w-full px-4 py-2 bg-white dark:bg-gray-800 border border-gray-300 dark:border-gray-600 rounded-full flex items-center justify-between hover:bg-gray-100 dark:hover:bg-gray-700 transition">
    <span id="selectedCategoryText" class="truncate">All Categories</span>
    <svg class="w-4 h-4 opacity-70" viewBox="0 0 20 20" fill="currentColor">
      <path d="M5.23 7.21a.75.75 0 011.06.02L10 10.94l3.71-3.71a.75.75 0 111.08 1.04l-4.25 4.25a.75.75 0 01-1.08 0L5.21 8.27a.75.75 0 01.02-1.06z"/>
    </svg>
  </button>

  <div id="categoryDropdownList"
     class="hidden absolute z-50 mt-2 w-full max-h-80 overflow-y-auto overscroll-contain scroll-smooth
            bg-white dark:bg-gray-800 border border-gray-200 dark:border-gray-700 rounded-lg shadow-xl">

    <!-- JS injects category items -->
  </div>
</div>
<!-- NEW: Level tab with Basic / Intermediate / Advanced buttons -->
<div id="level-tab" class="w-full">
  <div class="mt-2 flex flex-wrap items-centre justify-centre gap-2">
    <button id="btnLevelAll" role="button" tabindex="0"
            class="px-4 py-2 rounded-full border border-gray-300 dark:border-gray-600 text-sm sm:text-base">
      All
    </button>
    <button id="btnLevelBasic" role="button" tabindex="0"
            class="px-4 py-2 rounded-full border border-gray-300 dark:border-gray-600 text-sm sm:text-base">
      Basic
    </button>
    <button id="btnLevelIntermediate" role="button" tabindex="0"
            class="px-4 py-2 rounded-full border border-gray-300 dark:border-gray-600 text-sm sm:text-base">
      Intermediate
    </button>
    <button id="btnLevelAdvanced" role="button" tabindex="0"
            class="px-4 py-2 rounded-full border border-gray-300 dark:border-gray-600 text-sm sm:text-base">
      Advanced
    </button>
  </div>
</div>

<!-- Reset Button -->
<button id="btnResetFilters"
        class="ml-2 px-4 py-2 rounded-full border border-gray-300 dark:border-gray-600 text-sm sm:text-base hover:bg-gray-100 dark:hover:bg-gray-700">
  Reset All
</button>

            <div class="flex items-center gap-2">
                <button id="theme-toggle" class="p-2 rounded-full bg-gray-200 dark:bg-gray-700 hover:bg-gray-300 dark:hover:bg-gray-600 transition-colors">
                    <svg id="theme-icon-light" class="w-6 h-6 hidden" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 3v1m0 16v1m9-9h-1M4 12H3m15.364 6.364l-.707-.707M6.343 6.343l-.707-.707m12.728 0l-.707.707M6.343 17.657l-.707.707M16 12a4 4 0 11-8 0 4 4 0 018 0z"></path></svg>
                    <svg id="theme-icon-dark" class="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M20.354 15.354A9 9 0 018.646 3.646 9.003 9.003 0 0012 21a9.003 9.003 0 008.354-5.646z"></path></svg>
                </button>
                <!-- Menu Dropdown -->
                <div class="relative" id="menu-container">
                    <!-- Menu Button (Hamburger Icon) -->
                    <button id="menu-toggle-btn" class="p-2 rounded-full bg-gray-200 dark:bg-gray-700 hover:bg-gray-300 dark:hover:bg-gray-600 transition-colors" aria-haspopup="true" aria-expanded="false" aria-label="Main Menu">
                        <svg class="w-6 h-6" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke-width="1.5" stroke="currentColor">
                            <path stroke-linecap="round" stroke-linejoin="round" d="M3.75 6.75h16.5M3.75 12h16.5m-16.5 5.25h16.5" />
                        </svg>
                    </button>
                    <!-- Dropdown Panel -->
                    <div id="menu-dropdown" class="hidden absolute right-0 mt-2 w-48 bg-white dark:bg-gray-800 rounded-lg shadow-xl z-20 py-1" role="menu">
                        <a href="#" id="myProfileBtn" class="block px-4 py-2 text-sm text-gray-700 dark:text-gray-200 hover:bg-gray-100 dark:hover:bg-gray-700" role="menuitem">My Profile</a>
                        <a href="#" id="loginBtn" class="block px-4 py-2 text-sm text-gray-700 dark:text-gray-200 hover:bg-gray-100 dark:hover:bg-gray-700" role="menuitem">Login</a>
                        <a href="#" id="signupBtn" class="block px-4 py-2 text-sm text-gray-700 dark:text-gray-200 hover:bg-gray-100 dark:hover:bg-gray-700" role="menuitem">Sign Up</a>
                        <a href="#" id="logoutBtn" class="block px-4 py-2 text-sm text-gray-700 dark:text-gray-200 hover:bg-gray-100 dark:hover:bg-gray-700" role="menuitem">Logout</a>
                    </div>
                </div>
            </div>
        </div>

        <!-- Alphabet Filters -->
        <div id="alphabet-filters" class="flex flex-wrap justify-center gap-1 mb-6">
            <!-- Alphabet filter buttons will be injected here -->
        </div>

        <!-- Category Filters -->
        <div id="category-filters" class="hidden flex flex-wrap justify-center gap-2 mb-6">
            <!-- Filter buttons will be injected here -->
        </div>

        <!-- Content Grid -->
        <div id="content-grid" class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4 gap-6"></div>

        <!-- No Results Message -->
        <div id="no-results" class="text-center py-10 hidden">
            <h3 class="text-2xl font-semibold">No Results Found</h3>
            <p class="text-gray-500 dark:text-gray-400 mt-2">Try adjusting your search or filters.</p>
        </div>
    </div>
    
    <div id="copied-toast" class="copied-toast">Copied to clipboard!</div>

    <!-- Modals (Example, Auth, Profile, Welcome, Legal) -->
    <!-- Example Modal -->
    <div id="exampleModal" class="modal-base fixed inset-0 bg-black bg-opacity-60 z-50 hidden items-center justify-center p-4" role="dialog" aria-modal="true" aria-labelledby="modalTitle">
        <div id="modalContent" class="modal-content bg-white dark:bg-gray-800 rounded-lg shadow-xl w-full max-w-lg p-6 relative">
            <button id="closeModal" class="absolute top-3 right-3 text-gray-400 hover:text-gray-600 dark:hover:text-gray-200 transition-colors" aria-label="Close">
                <svg class="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M6 18L18 6M6 6l12 12"></path></svg>
            </button>
            <h3 id="modalTitle" class="text-2xl font-bold text-gray-900 dark:text-white mb-2"></h3>
            <p id="modalSyntax" class="font-mono text-sm bg-gray-100 dark:bg-gray-700 p-2 rounded break-words mb-4"></p>
            <p class="text-gray-700 dark:text-gray-300"><strong>Real-World Example:</strong></p>
            <p id="modalExample" class="text-gray-600 dark:text-gray-400 mt-1"></p>
        </div>
    </div>

    <!-- Login Modal -->
    <div id="loginModal" class="modal-base fixed inset-0 bg-black bg-opacity-60 z-50 hidden items-center justify-center p-4" role="dialog" aria-modal="true" aria-labelledby="loginModalTitle">
        <div class="modal-content bg-white dark:bg-gray-800 rounded-lg shadow-xl w-full max-w-sm p-6 relative">
            <button id="closeLoginModal" class="absolute top-3 right-3 text-gray-400 hover:text-gray-600 dark:hover:text-gray-200 transition-colors" aria-label="Close">
                <svg class="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M6 18L18 6M6 6l12 12"></path></svg>
            </button>
            <h3 id="loginModalTitle" class="text-2xl font-bold text-gray-900 dark:text-white mb-4">Login</h3>
            <div class="space-y-4">
                <input type="tel" id="loginMobileInput" placeholder="Mobile Number" class="w-full px-4 py-2 border rounded-lg bg-white dark:bg-gray-700 border-gray-300 dark:border-gray-600 focus:outline-none focus:ring-2 focus:ring-blue-500">
                <input type="password" id="loginPinInput" placeholder="4-Digit Entry PIN" pattern="[0-9]*" maxlength="4" inputmode="numeric" class="w-full px-4 py-2 border rounded-lg bg-white dark:bg-gray-700 border-gray-300 dark:border-gray-600 focus:outline-none focus:ring-2 focus:ring-blue-500">
                <p id="loginAuthError" class="text-red-500 text-sm text-center hidden"></p>
                <button id="submitLoginBtn" class="w-full px-4 py-2 bg-blue-500 text-white rounded-lg hover:bg-blue-600 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:ring-opacity-50 transition-colors">Login</button>
            </div>
        </div>
    </div>

    <!-- Sign Up Modal -->
    <div id="signupModal" class="modal-base fixed inset-0 bg-black bg-opacity-60 z-50 hidden items-center justify-center p-4" role="dialog" aria-modal="true" aria-labelledby="signupModalTitle">
        <div class="modal-content bg-white dark:bg-gray-800 rounded-lg shadow-xl w-full max-w-sm p-6 relative">
            <button id="closeSignupModal" class="absolute top-3 right-3 text-gray-400 hover:text-gray-600 dark:hover:text-gray-200 transition-colors" aria-label="Close">
                <svg class="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M6 18L18 6M6 6l12 12"></path></svg>
            </button>
            <h3 id="signupModalTitle" class="text-2xl font-bold text-gray-900 dark:text-white mb-4">Create Your Account</h3>
            <div class="space-y-3">
                <input type="text" id="signupNameInput" placeholder="Full Name" class="w-full px-4 py-2 border rounded-lg bg-white dark:bg-gray-700 border-gray-300 dark:border-gray-600 focus:outline-none focus:ring-2 focus:ring-blue-500">
                <input type="email" id="signupEmailInput" placeholder="Email Address" class="w-full px-4 py-2 border rounded-lg bg-white dark:bg-gray-700 border-gray-300 dark:border-gray-600 focus:outline-none focus:ring-2 focus:ring-blue-500">
                <input type="tel" id="signupMobileInput" placeholder="Mobile Number" class="w-full px-4 py-2 border rounded-lg bg-white dark:bg-gray-700 border-gray-300 dark:border-gray-600 focus:outline-none focus:ring-2 focus:ring-blue-500">
                <input type="text" id="signupPincodeInput" placeholder="Postal Code" class="w-full px-4 py-2 border rounded-lg bg-white dark:bg-gray-700 border-gray-300 dark:border-gray-600 focus:outline-none focus:ring-2 focus:ring-blue-500">
                <input type="password" id="signupPinInput" placeholder="Create 4-Digit Entry PIN" pattern="[0-9]*" maxlength="4" inputmode="numeric" class="w-full px-4 py-2 border rounded-lg bg-white dark:bg-gray-700 border-gray-300 dark:border-gray-600 focus:outline-none focus:ring-2 focus:ring-blue-500">
                <div class="flex items-center gap-2 pt-2">
                    <input type="checkbox" id="terms-agree-checkbox" class="h-4 w-4 rounded border-gray-300 text-blue-600 focus:ring-blue-500">
                    <label for="terms-agree-checkbox" class="text-xs text-gray-600 dark:text-gray-400">
                        I agree to the <a href="#" id="privacy-policy-link-signup" class="underline text-blue-500 hover:text-blue-600">Privacy Policy</a> and <a href="#" id="terms-of-use-link-signup" class="underline text-blue-500 hover:text-blue-600">Terms of Use</a>.
                    </label>
                </div>
                <p id="signupAuthError" class="text-red-500 text-sm text-center hidden"></p>
                <button id="submitSignupBtn" class="w-full px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:ring-opacity-50 transition-colors disabled:bg-gray-400 disabled:cursor-not-allowed" disabled>Sign Up</button>
            </div>
        </div>
    </div>

    <!-- My Profile Modal -->
    <div id="profileModal" class="modal-base fixed inset-0 bg-black bg-opacity-60 z-50 hidden items-center justify-center p-4" role="dialog" aria-modal="true" aria-labelledby="profileModalTitle">
        <div class="modal-content bg-white dark:bg-gray-800 rounded-lg shadow-xl w-full max-w-sm p-6 relative">
            <button id="closeProfileModal" class="absolute top-3 right-3 text-gray-400 hover:text-gray-600 dark:hover:text-gray-200 transition-colors" aria-label="Close">
                <svg class="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M6 18L18 6M6 6l12 12"></path></svg>
            </button>
            <h3 id="profileModalTitle" class="text-2xl font-bold text-gray-900 dark:text-white mb-4">My Profile</h3>
            <div id="profileDetails" class="space-y-2 text-gray-700 dark:text-gray-300">
                <!-- Profile details will be injected here -->
            </div>
        </div>
    </div>

    <!-- Welcome Modal -->
    <div id="welcomeModal" class="modal-base fixed inset-0 bg-black bg-opacity-60 z-50 hidden items-center justify-center p-4" role="dialog" aria-modal="true" aria-labelledby="welcomeMessage">
        <div class="modal-content bg-white dark:bg-gray-800 rounded-lg shadow-xl w-full max-w-sm p-6 text-center">
            <h3 id="welcomeMessage" class="text-2xl font-bold text-gray-900 dark:text-white mb-4"></h3>
            <p class="text-gray-600 dark:text-gray-400 mb-6">You're all set! Start exploring the world of Excel.</p>
            <button id="closeWelcomeModal" class="w-full px-4 py-2 bg-blue-500 text-white rounded-lg hover:bg-blue-600 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:ring-opacity-50 transition-colors">Get Started</button>
        </div>
    </div>

    <!-- Privacy Policy Modal -->
    <div id="privacyModal" class="modal-base fixed inset-0 bg-black bg-opacity-60 z-50 hidden items-center justify-center p-4" role="dialog" aria-modal="true" aria-labelledby="privacyModalTitle">
        <div class="modal-content bg-white dark:bg-gray-800 rounded-lg shadow-xl w-full max-w-2xl p-6 relative flex flex-col" style="max-height: 90vh;">
            <button id="closePrivacyModal" class="absolute top-3 right-3 text-gray-400 hover:text-gray-600 dark:hover:text-gray-200 transition-colors" aria-label="Close">
                <svg class="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M6 18L18 6M6 6l12 12"></path></svg>
            </button>
            <h3 id="privacyModalTitle" class="text-2xl font-bold text-gray-900 dark:text-white mb-4 flex-shrink-0">Privacy Policy</h3>
            <div class="policy-content overflow-y-auto pr-4 text-sm text-gray-600 dark:text-gray-300 space-y-4">
                <p class="italic text-xs">Last Updated: October 5, 2025</p>
                
                <div>
                    <h4 class="font-bold text-gray-800 dark:text-gray-200">1. Introduction</h4>
                    <p>Welcome to Excel Speed & Formula Vault ("we", "us", "our"). We are committed to protecting your personal information. This Privacy Policy explains how we collect, use, store, and disclose your information when you use our website.</p>
                </div>

                <div>
                    <h4 class="font-bold text-gray-800 dark:text-gray-200">2. Information We Collect</h4>
                    <p>We collect the following types of information:</p>
                    <ul class="list-disc list-inside ml-4">
                        <li><strong>Personal Information You Provide:</strong> When you create an account, we collect your Full Name, Email Address, Mobile Number, Postal Code, and a 4-digit Entry PIN ("Personal Data").</li>
                        <li><strong>Usage Information:</strong> We store your favorited formulas and shortcuts.</li>
                    </ul>
                </div>

                 <div>
                    <h4 class="font-bold text-gray-800 dark:text-gray-200">3. How We Use Your Information</h4>
                    <p>We use your information for the following purposes:</p>
                    <ul class="list-disc list-inside ml-4">
                        <li><strong>To Provide and Personalize Our Services:</strong> To create and manage your account, verify your identity for login, and display a personalized greeting.</li>
                        <li><strong>To Communicate With You:</strong> To send you a one-time welcome email upon registration and to respond to your inquiries through our contact form.</li>
                    </ul>
                </div>
                
                <div>
                    <h4 class="font-bold text-gray-800 dark:text-gray-200">4. Data Storage and Security</h4>
                    <p>All your Personal Data and usage information (like favorites) are stored exclusively in your web browser's `localStorage` on your own device. This means:</p>
                    <ul class="list-disc list-inside ml-4">
                        <li>Your data is **not** transmitted to or stored on any central server or cloud database owned by us.</li>
                        <li>We do not have access to the information you provide.</li>
                        <li>The security of your data depends on the security of your device and browser. For added security, your 4-digit PIN is not stored directly. Instead, we store a cryptographic hash of your PIN.</li>
                    </ul>
                </div>
                
                <div>
                    <h4 class="font-bold text-gray-800 dark:text-gray-200">5. Data Sharing and Disclosure</h4>
                    <p>We do not sell, rent, or trade your Personal Data. We only share your information in the following limited circumstance:</p>
                     <ul class="list-disc list-inside ml-4">
                        <li><strong>Third-Party Service Providers:</strong> We use EmailJS, a third-party service, to facilitate our contact form and send welcome emails. When you sign up or contact us, your name and email address are processed by EmailJS solely for the purpose of delivering these emails. We recommend you review the <a href="https://www.emailjs.com/legal/privacy-policy/" target="_blank" rel="noopener noreferrer" class="text-blue-500 underline">EmailJS Privacy Policy</a>.</li>
                    </ul>
                </div>

                <div>
                    <h4 class="font-bold text-gray-800 dark:text-gray-200">6. Your Data Rights and Choices</h4>
                    <p>In accordance with principles of data protection laws like the DPDP Act, you have rights over your data:</p>
                     <ul class="list-disc list-inside ml-4">
                        <li><strong>Right to Access:</strong> You can view the primary data we store (Name, Email, Mobile, Postal Code) at any time via the "My Profile" feature.</li>
                        <li><strong>Right to Erasure (Deletion):</strong> Since your data is stored locally on your device, you have full control. You can permanently delete all your account data and favorites by clearing your browser's cache and site data for this website.</li>
                        <li><strong>Logging Out:</strong> When you log out, your active session data is cleared, but your main account information and favorites remain in `localStorage` for your next visit unless you clear it as described above.</li>
                    </ul>
                </div>
                
                <div>
                    <h4 class="font-bold text-gray-800 dark:text-gray-200">7. Children's Privacy</h4>
                    <p>Our service is not intended for individuals under the age of 18. We do not knowingly collect personal information from children. If we become aware that we have inadvertently collected such information, we will take steps to delete it.</p>
                </div>
                
                <div>
                    <h4 class="font-bold text-gray-800 dark:text-gray-200">8. Changes to This Policy</h4>
                    <p>We may update this Privacy Policy from time to time. Any changes will be posted on this page with an updated revision date. We encourage you to review this policy periodically.</p>
                </div>
                
                <div>
                    <h4 class="font-bold text-gray-800 dark:text-gray-200">9. Contact Us</h4>
                    <p>If you have any questions about this Privacy Policy, please contact us using the form in the footer of this website.</p>
                </div>

            </div>
        </div>
    </div>
    
    <!-- Terms of Use Modal -->
    <div id="termsModal" class="modal-base fixed inset-0 bg-black bg-opacity-60 z-50 hidden items-center justify-center p-4" role="dialog" aria-modal="true" aria-labelledby="termsModalTitle">
        <div class="modal-content bg-white dark:bg-gray-800 rounded-lg shadow-xl w-full max-w-2xl p-6 relative flex flex-col" style="max-height: 90vh;">
            <button id="closeTermsModal" class="absolute top-3 right-3 text-gray-400 hover:text-gray-600 dark:hover:text-gray-200 transition-colors" aria-label="Close">
                <svg class="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M6 18L18 6M6 6l12 12"></path></svg>
            </button>
            <h3 id="termsModalTitle" class="text-2xl font-bold text-gray-900 dark:text-white mb-4 flex-shrink-0">Terms of Use</h3>
            <div class="policy-content overflow-y-auto pr-4 text-sm text-gray-600 dark:text-gray-300 space-y-4">
                <p class="italic text-xs">Last Updated: October 5, 2025</p>
                <p class="text-red-500 font-semibold text-xs">Legal Disclaimer: This is a sample Terms of Use document generated by an AI for informational purposes. It is NOT a substitute for professional legal advice. For a legally binding agreement, you must consult a qualified lawyer to tailor these terms to your specific circumstances and ensure compliance with local laws.</p>
                
                <div>
                    <h4 class="font-bold text-gray-800 dark:text-gray-200">1. Acceptance of Terms</h4>
                    <p>By accessing and using Excel Speed & Formula Vault (the "Website"), you accept and agree to be bound by these Terms of Use ("Terms") and our Privacy Policy. If you do not agree to these Terms, you must not access or use the Website.</p>
                </div>

                <div>
                    <h4 class="font-bold text-gray-800 dark:text-gray-200">2. Description of Service</h4>
                    <p>The Website provides a curated database of Microsoft Excel formulas and shortcuts, along with descriptions and examples, for educational and informational purposes only.</p>
                </div>

                <div>
                    <h4 class="font-bold text-gray-800 dark:text-gray-200">3. User Accounts</h4>
                    <p>To access certain features, you must register for an account. You agree to provide accurate, current, and complete information during registration. You are responsible for safeguarding your 4-digit PIN and for all activities that occur under your account. You agree to notify us immediately of any unauthorized use of your account.</p>
                </div>
                
                <div>
                    <h4 class="font-bold text-gray-800 dark:text-gray-200">4. Intellectual Property</h4>
                    <p>All content on this Website, including text, graphics, logos, and the compilation thereof, is the property of Excel Speed & Formula Vault or its content suppliers and is protected by copyright and other intellectual property laws. Microsoft® and Excel® are registered trademarks of Microsoft Corporation. This Website is not affiliated with, endorsed, or sponsored by Microsoft Corporation.</p>
                </div>

                <div>
                    <h4 class="font-bold text-gray-800 dark:text-gray-200">5. Disclaimer of Warranties</h4>
                    <p>The Website and its content are provided on an "as is" and "as available" basis without any warranties of any kind, either express or implied. We do not warrant that the information provided is accurate, reliable, complete, or current. Your use of the Website is at your sole risk.</p>
                </div>
                
                <div>
                    <h4 class="font-bold text-gray-800 dark:text-gray-200">6. Limitation of Liability</h4>
                    <p>To the fullest extent permitted by applicable law, in no event shall the owner of Excel Speed & Formula Vault, its affiliates, or their licensors, be liable for any indirect, incidental, special, consequential, or punitive damages, including but not limited to, loss of profits, data, use, goodwill, or other intangible losses, resulting from (i) your access to or use of or inability to access or use the Website; (ii) any conduct or content of any third party on the Website; (iii) any content obtained from the Website; and (iv) unauthorized access, use, or alteration of your transmissions or content, whether based on warranty, contract, tort (including negligence), or any other legal theory, whether or not we have been informed of the possibility of such damage.</p>
                </div>

                <div>
                    <h4 class="font-bold text-gray-800 dark:text-gray-200">7. Indemnification</h4>
                    <p>You agree to defend, indemnify, and hold harmless the owner of Excel Speed & Formula Vault and its licensee and licensors, and their employees, contractors, agents, officers, and directors, from and against any and all claims, damages, obligations, losses, liabilities, costs or debt, and expenses (including but not limited to attorney's fees), resulting from or arising out of a) your use and access of the Website, or b) a breach of these Terms.</p>
                </div>

                <div>
                    <h4 class="font-bold text-gray-800 dark:text-gray-200">8. Governing Law and Jurisdiction</h4>
                    <p>These Terms shall be governed and construed in accordance with the laws of India, without regard to its conflict of law provisions. You agree to submit to the exclusive jurisdiction of the courts located in Kolaghat, West Bengal, India to resolve any legal matter arising from these Terms.</p>
                </div>

                <div>
                    <h4 class="font-bold text-gray-800 dark:text-gray-200">9. Dispute Resolution</h4>
                    <p>Any dispute arising out of or in connection with these Terms shall first be attempted to be resolved through informal negotiations. If the dispute cannot be resolved informally, you agree that it will be resolved through binding arbitration in Kolaghat, West Bengal, under the Arbitration and Conciliation Act, 1996. The arbitration shall be conducted by a single arbitrator, and the language of the arbitration shall be English.</p>
                </div>

                <div>
                    <h4 class="font-bold text-gray-800 dark:text-gray-200">10. Class Action Waiver</h4>
                    <p>You agree that any arbitration or proceeding shall be limited to the dispute between us and you individually. To the full extent permitted by law, (i) no arbitration or proceeding shall be joined with any other; (ii) there is no right or authority for any dispute to be arbitrated or resolved on a class-action basis or to utilize class action procedures; and (iii) there is no right or authority for any dispute to be brought in a purported representative capacity on behalf of the general public or any other persons.</p>
                </div>

                <div>
                    <h4 class="font-bold text-gray-800 dark:text-gray-200">11. Changes to Terms</h4>
                    <p>We reserve the right, at our sole discretion, to modify or replace these Terms at any time. We will provide notice of any changes by posting the new Terms on this page. By continuing to access or use our Website after those revisions become effective, you agree to be bound by the revised terms.</p>
                </div>
            </div>
        </div>
    </div>

    <!-- Footer -->
    <footer class="mt-16 border-t border-gray-200 dark:border-gray-700 pt-8">
        <div class="container mx-auto px-4">
            <div class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-8">
                
                <!-- Founder Details -->
                <div class="text-center md:text-left">
                    <h3 class="text-xl font-bold mb-4">Founder</h3>
                    <div class="flex flex-col items-center md:items-start">
                        <img src="https://placehold.co/100x100/EBF4FF/3B82F6?text=AKD" alt="Founder Amit Kumar Dutta" class="rounded-full mb-3">
                        <p class="font-semibold text-lg">Amit Kumar Dutta</p>
                        <p class="text-sm text-gray-500 dark:text-gray-400">Excel Enthusiast & Developer</p>
                        <p class="text-xs text-gray-500 dark:text-gray-400 mt-2">Driven by a passion for efficiency, Amit created this vault to empower users of all levels to master Excel and unlock their data's full potential.</p>
                        <div class="flex space-x-4 mt-4">
                            <a href="#" class="text-gray-400 hover:text-blue-500"><svg class="w-6 h-6" fill="currentColor" viewBox="0 0 24 24"><path d="M19 0h-14c-2.761 0-5 2.239-5 5v14c0 2.761 2.239 5 5 5h14c2.762 0 5-2.239 5-5v-14c0-2.761-2.238-5-5-5zm-11 19h-3v-11h3v11zm-1.5-12.268c-.966 0-1.75-.79-1.75-1.764s.784-1.764 1.75-1.764 1.75.79 1.75 1.764-.783 1.764-1.75 1.764zm13.5 12.268h-3v-5.604c0-3.368-4-3.113-4 0v5.604h-3v-11h3v1.765c1.396-2.586 7-2.777 7 2.476v6.759z"/></svg></a>
                            <a href="#" class="text-gray-400 hover:text-blue-500"><svg class="w-6 h-6" fill="currentColor" viewBox="0 0 24 24"><path d="M24 4.557c-.883.392-1.832.656-2.828.775 1.017-.609 1.798-1.574 2.165-2.724-.951.564-2.005.974-3.127 1.195-.897-.957-2.178-1.555-3.594-1.555-3.179 0-5.515 2.966-4.797 6.045-4.091-.205-7.719-2.165-10.148-5.144-1.29 2.213-.669 5.108 1.523 6.574-.806-.026-1.566-.247-2.229-.616-.054 2.281 1.581 4.415 3.949 4.89-.693.188-1.452.232-2.224.084.626 1.956 2.444 3.379 4.6 3.419-2.07 1.623-4.678 2.588-7.52 2.588-1.259 0-2.397-.07-3.487-.206 2.618 1.68 5.743 2.662 9.14 2.662 10.97 0 17.001-9.033 16.63-17.396 .92-.665 1.713-1.49 2.345-2.433z"/></svg></a>
                        </div>
                    </div>
                </div>

                <!-- Contact Details -->
                <div class="md:col-span-2 lg:col-span-2">
                    <h3 class="text-xl font-bold mb-4">Contact Us</h3>
                    <p class="text-sm text-gray-500 dark:text-gray-400 mb-4">Have a question or suggestion? Drop a message!</p>
                    <form id="contact-form" class="space-y-4">
                        <div class="flex flex-col sm:flex-row gap-4">
                            <input type="text" id="contact_name" name="contact_name" placeholder="Your Name" class="w-full px-4 py-2 border rounded-lg bg-gray-200 dark:bg-gray-800 border-gray-300 dark:border-gray-600 focus:outline-none focus:ring-2 focus:ring-blue-500" required>
                            <input type="email" id="contact_email" name="contact_email" placeholder="Your Email" class="w-full px-4 py-2 border rounded-lg bg-gray-200 dark:bg-gray-800 border-gray-300 dark:border-gray-600 focus:outline-none focus:ring-2 focus:ring-blue-500" required>
                        </div>
                        <textarea id="contact_message" name="contact_message" placeholder="Your Message" rows="4" class="w-full px-4 py-2 border rounded-lg bg-gray-200 dark:bg-gray-800 border-gray-300 dark:border-gray-600 focus:outline-none focus:ring-2 focus:ring-blue-500" required></textarea>
                         <div class="flex items-center gap-2">
                            <input type="checkbox" id="contact-consent-checkbox" class="h-4 w-4 rounded border-gray-300 text-blue-600 focus:ring-blue-500">
                            <label for="contact-consent-checkbox" class="text-xs text-gray-600 dark:text-gray-400">
                                I consent to my name and email being sent via EmailJS. See their <a href="https://www.emailjs.com/legal/privacy-policy/" target="_blank" rel="noopener noreferrer" class="underline text-blue-500">privacy policy</a>.
                            </label>
                        </div>
                        <button type="submit" id="submit-contact-btn" class="px-6 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:ring-opacity-50 transition-colors disabled:bg-gray-400 disabled:cursor-not-allowed" disabled>Send Message</button>
                        <p id="contact-response" class="text-sm mt-2"></p>
                    </form>
                </div>

                 <!-- Testimonials -->
                 <div class="lg:col-span-1">
                    <h3 class="text-xl font-bold mb-4 text-center lg:text-left">Testimonials</h3>
                    <div class="space-y-4">
                        <div class="testimonial-card bg-white/50 dark:bg-gray-800/50 p-4 rounded-lg shadow-lg backdrop-blur-sm">
                            <p class="text-gray-600 dark:text-gray-300 italic">"This vault is a lifesaver! The search is so smart, I find exactly what I need in seconds."</p>
                            <div class="flex items-center mt-3">
                                <img src="https://placehold.co/40x40/EBF4FF/3B82F6?text=PS" alt="User" class="w-10 h-10 rounded-full mr-3">
                                <div>
                                    <p class="font-semibold">Priya Sharma</p>
                                    <p class="text-xs text-gray-500 dark:text-gray-400">Financial Analyst</p>
                                </div>
                            </div>
                        </div>
                        <div class="testimonial-card bg-white/50 dark:bg-gray-800/50 p-4 rounded-lg shadow-lg backdrop-blur-sm">
                            <p class="text-gray-600 dark:text-gray-300 italic">"The 3D floating cards and dark mode make learning Excel formulas surprisingly enjoyable. A+ design!"</p>
                            <div class="flex items-center mt-3">
                                <img src="https://placehold.co/40x40/EBF4FF/3B82F6?text=RK" alt="User" class="w-10 h-10 rounded-full mr-3">
                                <div>
                                    <p class="font-semibold">Rajesh Kumar</p>
                                    <p class="text-xs text-gray-500 dark:text-gray-400">Data Scientist</p>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
            
            <!-- Copyright & Legal Links -->
            <div class="mt-8 pt-4 border-t border-gray-300 dark:border-gray-700 text-center text-gray-500 dark:text-gray-400 text-sm">
                <p>&copy; <span id="copyright-year"></span> Excel Speed & Formula Vault. All Rights Reserved.</p>
                <div class="mt-2 space-x-4">
                    <a href="#" id="privacy-policy-link" class="hover:text-blue-500 underline">Privacy Policy</a>
                    <a href="#" id="terms-of-use-link" class="hover:text-blue-500 underline">Terms of Use</a>
                </div>
            </div>
        </div>
    </footer>


    <script>
        const formulas = [
            
  { name: 'EQUALS', category: 'Core Excel Starter Kit', description: 'Checks whether two values are exactly the same.', syntax: '=A1=B1', example: 'If A1 has ₹500 and B1 has ₹500, it returns TRUE. If B1 has ₹600, it returns FALSE.' },
  { name: 'PLUS', category: 'Core Excel Starter Kit', description: 'Adds two numbers or cell values.', syntax: '=A1+B1', example: 'If A1 is ₹1200 (salary) and B1 is ₹300 (allowance), result = ₹1500.' },
  { name: 'MINUS', category: 'Core Excel Starter Kit', description: 'Subtracts one value from another.', syntax: '=A1-B1', example: 'If A1 is ₹5000 (income) and B1 is ₹3200 (expense), result = ₹1800.' },
  { name: 'MULTIPLY', category: 'Core Excel Starter Kit', description: 'Multiplies two numbers.', syntax: '=A1*B1', example: 'If A1 is 45 units and B1 is ₹200 per unit, result = ₹9,000.' },
  { name: 'DIVIDE', category: 'Core Excel Starter Kit', description: 'Divides one number by another.', syntax: '=A1/B1', example: 'If total marks A1 = 450 and full marks B1 = 500, result = 0.9 (or 90%).' },
  { name: 'PERCENT', category: 'Core Excel Starter Kit', description: 'Converts a fraction to percentage.', syntax: '=A1*100 &"%"', example: 'If A1 = 0.25, it displays 25%.' },

  { name: 'SUM', category: 'Basics & Logic', description: 'Adds all numbers in a range.', syntax: '=SUM(B2:B10)', example: 'Adds total sales of all branches from B2:B10.' },
  { name: 'AVERAGE', category: 'Basics & Logic', description: 'Calculates the mean of selected numbers.', syntax: '=AVERAGE(C2:C6)', example: 'Finds average monthly expenses from January to May.' },
  { name: 'MIN', category: 'Basics & Logic', description: 'Finds the smallest number in a range.', syntax: '=MIN(D2:D8)', example: 'Returns the lowest electricity bill among all branches.' },
  { name: 'MAX', category: 'Basics & Logic', description: 'Finds the largest number in a range.', syntax: '=MAX(D2:D8)', example: 'Shows the highest loan disbursal in the district.' },
  { name: 'IF', category: 'Basics & Logic', description: 'Checks a condition and returns one value if true, another if false.', syntax: '=IF(E2>=75,"Eligible","Not Eligible")', example: 'If attendance in E2 is 80, shows "Eligible"; if 60, shows "Not Eligible".' },

  { name: 'FILTER', category: 'Array & Filter', description: 'Extracts records that meet a condition.', syntax: '=FILTER(A2:D50,D2:D50>100000)', example: 'Shows only customers whose loan amount exceeds ₹1,00,000.' },
  { name: 'UNIQUE', category: 'Array & Filter', description: 'Removes duplicate entries from a list.', syntax: '=UNIQUE(B2:B100)', example: 'Returns unique branch names from a branch list.' },
  { name: 'SORT', category: 'Array & Filter', description: 'Sorts data by column in ascending or descending order.', syntax: '=SORT(A2:C100,3,-1)', example: 'Sorts employee data by highest salary first.' },
  { name: 'SEQUENCE', category: 'Array & Filter', description: 'Generates a list of sequential numbers.', syntax: '=SEQUENCE(12)', example: 'Creates a numbered list from 1 to 12 for each month.' },
  { name: 'TRANSPOSE', category: 'Array & Filter', description: 'Switches rows and columns.', syntax: '=TRANSPOSE(A1:D1)', example: 'Converts 4-month headers in one row into a vertical list.' },

  { name: 'IFERROR', category: 'Error Handling', description: 'Displays a custom value if a formula gives an error.', syntax: '=IFERROR(A2/B2,0)', example: 'If B2 is blank, returns 0 instead of #DIV/0!.' },
  { name: 'ISERROR', category: 'Error Handling', description: 'Checks whether a formula returns any error.', syntax: '=ISERROR(C2)', example: 'Returns TRUE if C2 has #N/A or #VALUE!.' },
  { name: 'ISNUMBER', category: 'Error Handling', description: 'Checks if a cell contains a numeric value.', syntax: '=ISNUMBER(D2)', example: 'Returns TRUE if D2 has 1500 (loan amount).' },
  { name: 'ISTEXT', category: 'Error Handling', description: 'Checks if a cell contains text.', syntax: '=ISTEXT(E2)', example: 'Returns TRUE if E2 contains "Completed".' },
  { name: 'ISBLANK', category: 'Error Handling', description: 'Checks if a cell is empty.', syntax: '=ISBLANK(F2)', example: 'If F2 is blank, returns TRUE; if filled, FALSE.' },

  { name: 'DSUM', category: 'Database & Information', description: 'Sums a column from a table based on criteria.', syntax: '=DSUM(Data,"Amount",Criteria)', example: 'Sums all loan amounts where Branch = "Tamluk".' },
  { name: 'DCOUNT', category: 'Database & Information', description: 'Counts how many records meet a condition.', syntax: '=DCOUNT(Data,"Amount",Criteria)', example: 'Counts how many loans are above ₹2 lakh.' },
  { name: 'INFO', category: 'Database & Information', description: 'Gives system information such as current directory.', syntax: '=INFO("directory")', example: 'Shows the file folder path where the workbook is saved.' },
  { name: 'CELL', category: 'Database & Information', description: 'Gives details about a cell such as address or format.', syntax: '=CELL("address",A10)', example: 'Returns "$A$10".' },
  { name: 'GETPIVOTDATA', category: 'Database & Information', description: 'Extracts data from a PivotTable using field names.', syntax: '=GETPIVOTDATA("Sales",$A$3,"Region","East")', example: 'Fetches total sales for Region = East from Pivot Table.' },

  { name: 'XMATCH', category: 'Advanced Lookup', description: 'Finds the position of an item within a range.', syntax: '=XMATCH("East",A2:A50)', example: 'Returns the row number where "East" branch appears first.' },
  { name: 'CHOOSE', category: 'Advanced Lookup', description: 'Selects a value from a list using an index number.', syntax: '=CHOOSE(2,"Gold","Silver","Bronze")', example: 'Returns "Silver" as the 2nd choice.' },
  { name: 'OFFSET', category: 'Advanced Lookup', description: 'Moves reference by a certain number of rows and columns.', syntax: '=OFFSET(A1,2,3)', example: 'Refers to the cell two rows down and three columns right of A1.' },
  { name: 'ADDRESS', category: 'Advanced Lookup', description: 'Creates a cell address from row and column numbers.', syntax: '=ADDRESS(5,2)', example: 'Returns "$B$5".' },
  { name: 'INDIRECT', category: 'Advanced Lookup', description: 'Converts a text reference into an active reference.', syntax: '=INDIRECT("C"&A2)', example: 'If A2 = 10, formula refers to C10.' },

  { name: 'SUMIF', category: 'Conditional Math', description: 'Adds up values that meet one condition.', syntax: '=SUMIF(A2:A50,"East",C2:C50)', example: 'Sums all deposits for East region branches.' },
  { name: 'COUNTIF', category: 'Conditional Math', description: 'Counts how many cells meet a condition.', syntax: '=COUNTIF(B2:B50,">5000")', example: 'Counts how many accounts have balance above ₹5000.' },
  { name: 'AVERAGEIF', category: 'Conditional Math', description: 'Calculates average only for cells matching criteria.', syntax: '=AVERAGEIF(D2:D50,"Completed",E2:E50)', example: 'Average loan amount for records marked "Completed".' },
  { name: 'IFS', category: 'Conditional Math', description: 'Evaluates multiple conditions in order.', syntax: '=IFS(F2>=90,"A+",F2>=80,"A",F2>=70,"B","C")', example: 'Assigns grade to student based on percentage.' },
  { name: 'SWITCH', category: 'Conditional Math', description: 'Matches a value against multiple options and returns the result.', syntax: '=SWITCH(G2,"M","Male","F","Female","Other")', example: 'If G2 is F, returns "Female".' },

  { name: 'SPARKLINE', category: 'Dashboard & Sheet Info', description: 'Draws a small chart inside one cell.', syntax: '=SPARKLINE(B2:B13,"line")', example: 'Displays a trend of monthly savings from Jan to Dec.' },
  { name: 'HYPERLINK', category: 'Dashboard & Sheet Info', description: 'Creates clickable links to websites or sheets.', syntax: '=HYPERLINK("https://pnbindia.in","Visit PNB")', example: 'Click to open PNB official website.' },
  { name: 'FORMULATEXT', category: 'Dashboard & Sheet Info', description: 'Shows the actual formula from another cell.', syntax: '=FORMULATEXT(C5)', example: 'If C5 has =SUM(B2:B10), this displays that text.' },
  { name: 'SHEET', category: 'Dashboard & Sheet Info', description: 'Returns the sheet number.', syntax: '=SHEET()', example: 'If used in Sheet2, result is 2.' },
  { name: 'SHEETS', category: 'Dashboard & Sheet Info', description: 'Counts all sheets in a workbook.', syntax: '=SHEETS()', example: 'If there are 5 worksheets, returns 5.' },

  { name: 'MAP', category: 'Dynamic Array & Lambda', description: 'Applies a formula to every element in a range.', syntax: '=MAP(A2:A6,LAMBDA(x,x*1.1))', example: 'Increases all sales figures by 10%.' },
  { name: 'REDUCE', category: 'Dynamic Array & Lambda', description: 'Combines a range into one result using a running calculation.', syntax: '=REDUCE(0,A2:A10,LAMBDA(a,b,a+b))', example: 'Adds all transaction values step by step.' },
  { name: 'BYROW', category: 'Dynamic Array & Lambda', description: 'Performs a function on each row and spills the results.', syntax: '=BYROW(B2:D10,LAMBDA(r,SUM(r)))', example: 'Calculates each employee’s total score from columns B–D.' },
  { name: 'LAMBDA', category: 'Dynamic Array & Lambda', description: 'Creates a custom function inside Excel.', syntax: '=LAMBDA(x,x^2)(5)', example: 'Returns 25 by squaring 5.' },
  { name: 'LET', category: 'Dynamic Array & Lambda', description: 'Assigns a name to part of a formula for reusability.', syntax: '=LET(x,SUM(B2:B6),x/COUNT(B2:B6))', example: 'Finds average sales without repeating SUM or COUNT.' },

  { name: 'SUM', category: 'Basic Math & Logical Checks', description: 'Adds up numeric values.', syntax: '=SUM(C2:C10)', example: 'Calculates total of monthly EMI payments.' },
  { name: 'AVERAGE', category: 'Basic Math & Logical Checks', description: 'Calculates average of numbers.', syntax: '=AVERAGE(C2:C10)', example: 'Finds average electricity bill for 9 months.' },
  { name: 'MIN', category: 'Basic Math & Logical Checks', description: 'Finds smallest number.', syntax: '=MIN(D2:D12)', example: 'Finds lowest daily withdrawal amount.' },
  { name: 'MAX', category: 'Basic Math & Logical Checks', description: 'Finds largest number.', syntax: '=MAX(D2:D12)', example: 'Shows highest credit card bill paid in the year.' },
  { name: 'IF', category: 'Basic Math & Logical Checks', description: 'Returns a value based on a condition.', syntax: '=IF(E2>50000,"High Value","Normal")', example: 'If E2 = ₹60,000 deposit, it displays "High Value".' },
  { name: 'AND', category: 'Basic Math & Logical Checks', description: 'Tests if all conditions are TRUE.', syntax: '=AND(F2>0,F2<100)', example: 'Checks if discount in F2 is between 0 and 100%.' },
  { name: 'OR', category: 'Basic Math & Logical Checks', description: 'Tests if any condition is TRUE.', syntax: '=OR(G2="Yes",H2="Approved")', example: 'Returns TRUE if loan is either marked Yes or Approved.' },

  { name: 'TODAY', category: 'Date & Time', description: 'Returns the current system date.', syntax: '=TODAY()', example: 'If today is 6-Oct-2025, it returns 06-10-2025.' },
  { name: 'NOW', category: 'Date & Time', description: 'Returns current date and time.', syntax: '=NOW()', example: 'Shows 06-10-2025 08:45 PM.' },
  { name: 'DATE', category: 'Date & Time', description: 'Combines year, month, and day into a date.', syntax: '=DATE(2025,10,6)', example: 'Displays 06-10-2025.' },
  { name: 'YEAR', category: 'Date & Time', description: 'Extracts the year from a date.', syntax: '=YEAR(A2)', example: 'If A2 has 12-Aug-2023, it returns 2023.' },
  { name: 'MONTH', category: 'Date & Time', description: 'Extracts the month number from a date.', syntax: '=MONTH(A2)', example: 'If A2 = 15-Sep-2024, it returns 9.' },
  { name: 'DAY', category: 'Date & Time', description: 'Extracts the day number from a date.', syntax: '=DAY(A2)', example: 'If A2 = 15-Sep-2024, it returns 15.' },
  { name: 'WEEKDAY', category: 'Date & Time', description: 'Returns the day of week (1=Sunday…7=Saturday).', syntax: '=WEEKDAY(A2)', example: 'If A2 = 07-Oct-2025, it returns 3 (Tuesday).' },
  { name: 'EOMONTH', category: 'Date & Time', description: 'Returns the last day of a month after a given offset.', syntax: '=EOMONTH(A2,0)', example: 'If A2 = 10-Sep-2025, returns 30-Sep-2025.' },
  { name: 'DATEDIF', category: 'Date & Time', description: 'Calculates difference between two dates.', syntax: '=DATEDIF(A2,B2,"D")', example: 'If A2=01-Jan-2025, B2=31-Jan-2025, result=30 days.' },
  { name: 'NETWORKDAYS', category: 'Date & Time', description: 'Counts working days between two dates excluding weekends.', syntax: '=NETWORKDAYS(A2,B2)', example: 'Between 1-Oct-2025 and 10-Oct-2025 → 8 working days.' },

  { name: 'PMT', category: 'Financial & Banking', description: 'Calculates EMI for a loan based on rate and tenure.', syntax: '=PMT(10%/12,60,-500000)', example: 'For ₹5 L loan @10% for 5 yrs, EMI = ₹10,624.' },
  { name: 'FV', category: 'Financial & Banking', description: 'Returns future value of an investment.', syntax: '=FV(0.08,10,0,-100000)', example: '₹1 L in FD for 10 yrs @8% → ₹2,15,892.' },
  { name: 'NPV', category: 'Financial & Banking', description: 'Calculates net present value of future cash flows.', syntax: '=NPV(0.1,B2:B6)+B1', example: 'Evaluates NPV of project with 10% discount rate.' },
  { name: 'IRR', category: 'Financial & Banking', description: 'Finds internal rate of return for cash flows.', syntax: '=IRR(A1:A6)', example: 'If investment and returns listed in A1:A6, gives yearly ROI%.' },
  { name: 'RATE', category: 'Financial & Banking', description: 'Finds interest rate per period.', syntax: '=RATE(60,-10624,500000)*12', example: 'Returns 10% for a ₹5 L loan paid in ₹10,624 EMIs for 5 yrs.' },

  { name: 'CUMIPMT', category: 'Loan & Credit', description: 'Total interest paid over a period.', syntax: '=CUMIPMT(10%/12,60,500000,1,12,0)', example: 'Shows total interest paid in first 12 months of a ₹5 L loan.' },
  { name: 'CUMPRINC', category: 'Loan & Credit', description: 'Total principal repaid over a period.', syntax: '=CUMPRINC(10%/12,60,500000,1,12,0)', example: 'Shows total principal repaid in first year.' },
  { name: 'IPMT', category: 'Loan & Credit', description: 'Interest part of EMI for a specific month.', syntax: '=IPMT(10%/12,1,60,-500000)', example: 'Interest in first EMI ≈ ₹4,167.' },
  { name: 'PPMT', category: 'Loan & Credit', description: 'Principal part of EMI for a specific month.', syntax: '=PPMT(10%/12,1,60,-500000)', example: 'Principal in first EMI ≈ ₹6,457.' },
  { name: 'DB', category: 'Loan & Credit', description: 'Depreciation using fixed-declining balance method.', syntax: '=DB(500000,50000,5,1)', example: 'Depreciation of ₹5 L asset in 1st yr to ₹4.5 L.' },


  { name: 'CAGR', category: 'Investment & Returns', description: 'Compound annual growth rate over years.', syntax: '=((End/Start)^(1/Years))-1', example: 'From ₹2 L to ₹4 L in 5 yrs → 14.9% CAGR.' },
  { name: 'ROI', category: 'Investment & Returns', description: 'Simple return on investment.', syntax: '=(Gain-Cost)/Cost', example: 'Bought at ₹1 L, sold at ₹1.3 L → 30% ROI.' },
  { name: 'XIRR', category: 'Investment & Returns', description: 'Annualized ROI for irregular cashflows.', syntax: '=XIRR(Values,Dates)', example: 'Used for SIP investments to find true annual return.' },
  { name: 'NPER', category: 'Investment & Returns', description: 'Number of periods for an investment or loan.', syntax: '=NPER(10%/12,-10624,500000)', example: 'Gives ~60 months for a ₹5 L loan at ₹10,624 EMI.' },
  { name: 'PV', category: 'Investment & Returns', description: 'Present value of a future sum or cashflow.', syntax: '=PV(0.08,10,0,-100000)', example: 'Present value of ₹1 L receivable in 10 yrs @8% = ₹46,319.' },


  { name: 'MEDIAN', category: 'Statistical', description: 'Middle value in a dataset.', syntax: '=MEDIAN(B2:B10)', example: 'If salaries range ₹25k–₹80k, median might be ₹45k.' },
  { name: 'MODE', category: 'Statistical', description: 'Most frequently occurring number.', syntax: '=MODE(C2:C20)', example: 'Shows the most common transaction amount.' },
  { name: 'STDEV', category: 'Statistical', description: 'Measures variation from mean.', syntax: '=STDEV(D2:D12)', example: 'Checks how consistent monthly sales are.' },
  { name: 'VAR', category: 'Statistical', description: 'Variance — spread of data.', syntax: '=VAR(D2:D12)', example: 'Shows how widely branch targets vary.' },
  { name: 'CORREL', category: 'Statistical', description: 'Correlation between two datasets.', syntax: '=CORREL(E2:E10,F2:F10)', example: 'Measures relation between deposits and profit.' },


  { name: 'CONCAT', category: 'Text & Formatting', description: 'Joins text strings together.', syntax: '=CONCAT(A2," ",B2)', example: 'Combines First Name and Last Name into full name.' },
  { name: 'TEXTJOIN', category: 'Text & Formatting', description: 'Joins text with a separator.', syntax: '=TEXTJOIN(", ",TRUE,A2:A6)', example: 'Combines branch names as "Kolaghat, Tamluk, Haldia".' },
  { name: 'UPPER', category: 'Text & Formatting', description: 'Converts text to uppercase.', syntax: '=UPPER(A2)', example: 'If A2 = "kolaghat", it returns "KOLAGHAT".' },
  { name: 'LOWER', category: 'Text & Formatting', description: 'Converts text to lowercase.', syntax: '=LOWER(A2)', example: 'If A2 = "PNB INDIA", returns "pnb india".' },
  { name: 'PROPER', category: 'Text & Formatting', description: 'Capitalizes first letter of each word.', syntax: '=PROPER(A2)', example: 'If A2 = "pnb purba medinipur", returns "Pnb Purba Medinipur".' },


  { name: 'COUNT', category: 'Count & Frequency', description: 'Counts cells with numbers.', syntax: '=COUNT(B2:B50)', example: 'Counts how many customers have numeric IDs entered.' },
  { name: 'COUNTA', category: 'Count & Frequency', description: 'Counts all non-blank cells.', syntax: '=COUNTA(A2:A20)', example: 'Counts total filled cells including text and numbers.' },
  { name: 'COUNTBLANK', category: 'Count & Frequency', description: 'Counts empty cells in a range.', syntax: '=COUNTBLANK(C2:C20)', example: 'Shows number of customers with missing phone numbers.' },
  { name: 'COUNTIFS', category: 'Count & Frequency', description: 'Counts cells meeting multiple criteria.', syntax: '=COUNTIFS(A2:A50,"East",B2:B50,">50000")', example: 'Counts East-region accounts above ₹50,000 balance.' },
  { name: 'FREQUENCY', category: 'Count & Frequency', description: 'Returns frequency distribution.', syntax: '=FREQUENCY(B2:B100,{10000,50000,100000})', example: 'Counts how many deposits fall into these ranges.' },


  { name: 'ROUND', category: 'Math & Rounding', description: 'Rounds a number to a set number of decimals.', syntax: '=ROUND(A2,2)', example: 'If A2 = 123.4567, result = 123.46.' },
  { name: 'ROUNDUP', category: 'Math & Rounding', description: 'Rounds a number upward.', syntax: '=ROUNDUP(A2,0)', example: 'If A2 = 12.1, result = 13.' },
  { name: 'ROUNDDOWN', category: 'Math & Rounding', description: 'Rounds a number downward.', syntax: '=ROUNDDOWN(A2,0)', example: 'If A2 = 12.9, result = 12.' },
  { name: 'INT', category: 'Math & Rounding', description: 'Removes decimal portion of a number.', syntax: '=INT(B2)', example: 'If B2 = 99.99, result = 99.' },
  { name: 'MOD', category: 'Math & Rounding', description: 'Returns remainder after division.', syntax: '=MOD(25,7)', example: '25 ÷ 7 leaves remainder 4.' },

  { name: 'NOT', category: 'Logical & Comparison', description: 'Reverses a logical value.', syntax: '=NOT(A2="Closed")', example: 'TRUE if account is not Closed.' },
  { name: 'XOR', category: 'Logical & Comparison', description: 'TRUE if one condition is true, not both.', syntax: '=XOR(A2="Yes",B2="Yes")', example: 'TRUE if exactly one of the two approvals is Yes.' },
  { name: 'IFNA', category: 'Logical & Comparison', description: 'Returns custom value if formula gives #N/A.', syntax: '=IFNA(VLOOKUP(A2,Data,2,0),"Not Found")', example: 'Shows "Not Found" if customer ID missing.' },


  { name: 'VLOOKUP', category: 'Lookup & Reference', description: 'Searches for a value in first column and returns data from another column.', syntax: '=VLOOKUP(A2,Table,3,0)', example: 'Looks up Customer ID in A2 and returns Mobile Number from 3rd column.' },
  { name: 'HLOOKUP', category: 'Lookup & Reference', description: 'Searches across top row and returns data from a row below.', syntax: '=HLOOKUP("Mar",A1:M4,3,0)', example: 'Finds March sales in the 3rd row of table A1:M4.' },
  { name: 'INDEX', category: 'Lookup & Reference', description: 'Returns value at given row and column.', syntax: '=INDEX(B2:D10,3,2)', example: 'Returns value from 3rd row, 2nd column of range B2:D10.' },
  { name: 'MATCH', category: 'Lookup & Reference', description: 'Returns position of a value in a range.', syntax: '=MATCH("East",A2:A10,0)', example: 'Returns row number of East branch in list.' },

  { name: 'STDEV.S', category: 'Probability & Statistical Advanced', description: 'Sample standard deviation — how spread out values are.', syntax: '=STDEV.S(B2:B25)', example: 'Measure variation in daily cash deposits across 24 days.' },
  { name: 'VAR.S', category: 'Probability & Statistical Advanced', description: 'Sample variance — square of standard deviation.', syntax: '=VAR.S(B2:B25)', example: 'Compare branch-to-branch variability in monthly sales.' },
  { name: 'CORREL', category: 'Probability & Statistical Advanced', description: 'Correlation between two datasets.', syntax: '=CORREL(C2:C25,D2:D25)', example: 'Relationship between advertising spend (C) and new accounts opened (D).' },
  { name: 'FORECAST.LINEAR', category: 'Probability & Statistical Advanced', description: 'Predicts a future value using linear trend.', syntax: '=FORECAST.LINEAR(2026, Y_Year, X_Year)', example: 'Predict 2026 annual deposits using 5 years of history.' },
  { name: 'NORM.DIST', category: 'Probability & Statistical Advanced', description: 'Normal distribution probability or cumulative.', syntax: '=NORM.DIST(x,mean,sd,TRUE)', example: 'Find probability a customer’s monthly spend is below ₹20,000 given mean & SD.' },
  { name: 'NORM.S.INV', category: 'Probability & Statistical Advanced', description: 'Inverse of standard normal (Z) distribution.', syntax: '=NORM.S.INV(0.975)', example: 'Returns ≈1.96 for 95% confidence interval calculations.' },
  { name: 'RANDARRAY', category: 'Probability & Statistical Advanced', description: 'Generates an array of random numbers (optionally integers).', syntax: '=RANDARRAY(5,1,1,100,TRUE)', example: 'Create 5 random test amounts between ₹1 and ₹100 for demo dashboards.' },


  { name: 'EXACT', category: 'Logical & Audit Helper', description: 'Case-sensitive comparison of two text values.', syntax: '=EXACT(A2,B2)', example: 'Validate IFSC codes typed in two cells match exactly.' },
  { name: 'XOR', category: 'Logical & Audit Helper', description: 'TRUE if exactly one condition is TRUE (not both).', syntax: '=XOR(C2="Approved",D2="Approved")', example: 'Flag if approval came from either Manager or Auditor, but not both.' },
  { name: 'TRUE', category: 'Logical & Audit Helper', description: 'Returns the logical value TRUE.', syntax: '=TRUE()', example: 'Used as a constant in validation formulas or tests.' },
  { name: 'FALSE', category: 'Logical & Audit Helper', description: 'Returns the logical value FALSE.', syntax: '=FALSE()', example: 'Used to lock a formula outcome while testing conditions.' },
  { name: 'FORMULATEXT', category: 'Logical & Audit Helper', description: 'Shows the actual formula from a referenced cell.', syntax: '=FORMULATEXT(E5)', example: 'Audit: Display the formula used in KPI cell E5 for review.' },


  { name: 'EMI (PMT)', category: 'Custom Financial Utilities', description: 'Monthly installment for loan based on rate and tenure.', syntax: '=PMT(9.5%/12,240,-1500000)', example: '₹15 L home loan, 20 years @ 9.5% → monthly EMI ≈ ₹13,985.' },
  { name: 'Interest Component (IPMT)', category: 'Custom Financial Utilities', description: 'Interest portion of a specific EMI.', syntax: '=IPMT(9.5%/12,1,240,-1500000)', example: 'First month interest on the above loan ≈ ₹11,875.' },
  { name: 'Principal Component (PPMT)', category: 'Custom Financial Utilities', description: 'Principal portion of a specific EMI.', syntax: '=PPMT(9.5%/12,1,240,-1500000)', example: 'First month principal ≈ ₹2,110.' },
  { name: 'Future Value (FV)', category: 'Custom Financial Utilities', description: 'Future value of a fixed deposit or investment.', syntax: '=FV(0.07,5,0,-100000)', example: '₹1,00,000 FD @7% for 5 years → ≈ ₹1,40,255.' },
  { name: 'CAGR (manual)', category: 'Custom Financial Utilities', description: 'Compound annual growth rate between start and end values.', syntax: '=((End/Start)^(1/Years))-1', example: '₹5,00,000 → ₹8,00,000 in 4 years → CAGR ≈ 12.9%.' },
  { name: 'Simple Interest (manual)', category: 'Custom Financial Utilities', description: 'Simple interest calculation for quick checks.', syntax: '=Principal*Rate*Time', example: '₹2,00,000 @ 8% for 1 year → ₹16,000 interest.' },


  { name: 'COVARIANCE.S', category: 'Advanced Statistical & Analytical Tools', description: 'Sample covariance — direction of joint movement.', syntax: '=COVARIANCE.S(B2:B25,C2:C25)', example: 'See if daily deposits and daily withdrawals move together.' },
  { name: 'COVARIANCE.P', category: 'Advanced Statistical & Analytical Tools', description: 'Population covariance for full datasets.', syntax: '=COVARIANCE.P(B2:B365,C2:C365)', example: 'Annual full data: deposits vs footfall covariance.' },
  { name: 'FORECAST.ETS', category: 'Advanced Statistical & Analytical Tools', description: 'Forecasts using exponential smoothing (seasonality-aware).', syntax: '=FORECAST.ETS(A13,A2:A12,B2:B12,1,1)', example: 'Predict next month’s credit card spends with monthly seasonality.' },
  { name: 'FORECAST.ETS.SEASONALITY', category: 'Advanced Statistical & Analytical Tools', description: 'Detects the seasonal cycle length.', syntax: '=FORECAST.ETS.SEASONALITY(A2:A36,B2:B36,1,1)', example: 'Identify a 12-month seasonal pattern in deposits.' },
  { name: 'LINEST', category: 'Advanced Statistical & Analytical Tools', description: 'Linear regression coefficients (slope/intercept, etc.).', syntax: '=LINEST(Y_range,X_range,TRUE,TRUE)', example: 'Estimate how sales (Y) change with ad spend (X).' },
  { name: 'LOGEST', category: 'Advanced Statistical & Analytical Tools', description: 'Exponential growth model fitting.', syntax: '=LOGEST(Y_range,X_range)', example: 'Model compounding user growth from marketing campaigns.' },
  { name: 'RSQ', category: 'Advanced Statistical & Analytical Tools', description: 'R² — goodness of fit for linear regression.', syntax: '=RSQ(Y_range,X_range)', example: 'Check how well ad spend explains sales (0–1 scale).' },
  { name: 'TREND', category: 'Advanced Statistical & Analytical Tools', description: 'Returns values along a linear trend.', syntax: '=TREND(Y_range,X_range,New_Xs)', example: 'Extend revenue trend to next 6 months.' },
  { name: 'GROWTH', category: 'Advanced Statistical & Analytical Tools', description: 'Returns values along an exponential curve.', syntax: '=GROWTH(Y_range,X_range,New_Xs)', example: 'Project app downloads assuming exponential growth.' },


  { name: 'MMULT', category: 'Array Math & Matrix Operations', description: 'Matrix multiplication of two arrays.', syntax: '=MMULT(A2:B3,D2:E3)', example: 'Score weights × scores to compute weighted totals.' },
  { name: 'MINVERSE', category: 'Array Math & Matrix Operations', description: 'Inverse of a square matrix.', syntax: '=MINVERSE(A2:D5)', example: 'Solve small systems of linear equations for planning models.' },
  { name: 'MDETERM', category: 'Array Math & Matrix Operations', description: 'Determinant of a matrix.', syntax: '=MDETERM(A2:D5)', example: 'Check if a coefficient matrix is invertible (non-zero determinant).' },
  { name: 'TRANSPOSE', category: 'Array Math & Matrix Operations', description: 'Swap rows and columns of a range.', syntax: '=TRANSPOSE(A1:D2)', example: 'Rotate a 2-row product list into a 2-column list for printing.' },
  { name: 'SUMPRODUCT', category: 'Array Math & Matrix Operations', description: 'Sum of products of corresponding array elements.', syntax: '=SUMPRODUCT(Quantity, Price)', example: 'Invoice total = sum of (Qty × Unit Price) across lines.' },
  { name: 'FREQUENCY', category: 'Array Math & Matrix Operations', description: 'Distribution counts across bins.', syntax: '=FREQUENCY(Sales,{10000,50000,100000})', example: 'Count how many transactions fall under ₹10k, ₹50k, and ₹1L+' },


  { name: 'TEXTBEFORE', category: 'Dynamic Data Cleaning & Transformation', description: 'Returns text before a specific delimiter.', syntax: '=TEXTBEFORE(A2,",")', example: 'From "Kolkata, West Bengal" → "Kolkata".' },
  { name: 'TEXTAFTER', category: 'Dynamic Data Cleaning & Transformation', description: 'Returns text after a specific delimiter.', syntax: '=TEXTAFTER(A2,"-")', example: 'From "ACC-9876" → "9876".' },
  { name: 'TEXTSPLIT + TRANSPOSE', category: 'Dynamic Data Cleaning & Transformation', description: 'Split a delimited string and stack it vertically.', syntax: '=TRANSPOSE(TEXTSPLIT(A2,","))', example: 'Turn "Gold,Silver,Bronze" into a 3-row list for validation.' },
  { name: 'TRIM with NBSP fix', category: 'Dynamic Data Cleaning & Transformation', description: 'Removes extra and non-breaking spaces from pasted web text.', syntax: '=TRIM(SUBSTITUTE(A2,CHAR(160),""))', example: 'Clean customer names pasted from web forms.' },
  { name: 'REGEX.REPLACE', category: 'Dynamic Data Cleaning & Transformation', description: 'Replace characters that match a pattern.', syntax: '=REGEX.REPLACE(A2,"[^0-9]","")', example: 'Strip all non-digits from a phone number field.' },
  { name: 'REGEX.EXTRACT', category: 'Dynamic Data Cleaning & Transformation', description: 'Extract substring that matches a pattern.', syntax: '=REGEX.EXTRACT(A2,"[A-Z]{3}[0-9]{4}")', example: 'Extract a PAN-like code such as ABC1234 from mixed text.' },
  { name: 'REGEX.MATCH', category: 'Dynamic Data Cleaning & Transformation', description: 'Check if text matches a pattern (TRUE/FALSE).', syntax: '=REGEX.MATCH(A2,"^[0-9]{6}$")', example: 'Validate that a pincode is exactly 6 digits.' },


  { name: 'Month Start', category: 'Business Calendar Automation', description: 'First day of the month for a given date.', syntax: '=EOMONTH(A2,-1)+1', example: 'If A2 = 15-May-2025 → 01-May-2025 (for monthly reports).' },
  { name: 'Quarter Label', category: 'Business Calendar Automation', description: 'Returns Q1–Q4 based on date.', syntax: '="Q"&INT((MONTH(A2)+2)/3)', example: 'If A2 = 10-Aug-2025 → "Q3".' },
  { name: 'Fiscal Year (Apr–Mar)', category: 'Business Calendar Automation', description: 'Indian-style FY start/end.', syntax: '=IF(MONTH(A2)>=4,YEAR(A2),YEAR(A2)-1)', example: 'If A2 = 10-Feb-2025 → FY 2024–25.' },
  { name: 'Days in Month', category: 'Business Calendar Automation', description: 'How many days in the date’s month.', syntax: '=DAY(EOMONTH(A2,0))', example: 'For Feb 2024 → 29 days (leap year).' },
  { name: 'Last Working Day', category: 'Business Calendar Automation', description: 'Last weekday of the current month.', syntax: '=WORKDAY(EOMONTH(A2,1)+1,-1)', example: 'Helpful for payroll cut-off dates.' },
  { name: 'Next EMI Date', category: 'Business Calendar Automation', description: 'Month-end of next month.', syntax: '=EOMONTH(TODAY(),1)', example: 'Set reminders for upcoming EMI collections.' },


  { name: 'NPV', category: 'Financial Modelling', description: 'Net present value of cash flows at a discount rate.', syntax: '=NPV(0.12,B2:B6)+B1', example: 'Project cash inflows in B2:B6; B1 is initial outflow (negative). Returns NPV at 12%.' },
  { name: 'IRR', category: 'Financial Modelling', description: 'Internal rate of return for periodic cash flows.', syntax: '=IRR(B1:B6)', example: 'Given -₹10L (B1) followed by returns in B2:B6, returns the project IRR.' },
  { name: 'MIRR', category: 'Financial Modelling', description: 'Modified IRR using finance and reinvest rates.', syntax: '=MIRR(B1:B6,0.10,0.08)', example: 'Assume 10% cost of capital and 8% reinvestment rate for realism.' },
  { name: 'XIRR', category: 'Financial Modelling', description: 'IRR for irregularly timed cash flows.', syntax: '=XIRR(Values,Dates)', example: 'Use for SIPs where deposits happen on different dates.' },
  { name: 'XNPV', category: 'Financial Modelling', description: 'NPV for irregular cash flows aligned to actual dates.', syntax: '=XNPV(0.12,Values,Dates)', example: 'Discount investor payouts using exact transaction dates at 12%.' },
  { name: 'ROI', category: 'Financial Modelling', description: 'Simple return on investment.', syntax: '=(Gain-Cost)/Cost', example: 'Buy at ₹1,00,000 and sell at ₹1,25,000 → ROI = 25%.' },
  { name: 'CAGR', category: 'Financial Modelling', description: 'Average annual growth rate across years.', syntax: '=((End/Start)^(1/Years))-1', example: '₹6L → ₹9L in 3 years → CAGR ≈ 14.5%.' },
  { name: 'NORM.INV', category: 'Financial Modelling', description: 'Value at a given probability for a normal distribution.', syntax: '=NORM.INV(0.95,Mean,SD)', example: 'Estimate 95th percentile annual return given mean & SD.' },


  { name: 'CUBESET', category: 'Power Query / Data Model Helpers', description: 'Defines a set from an OLAP cube or Data Model for later querying.' }, 
  { name: 'Deposit Growth %', category: 'Bank Branch Performance Metrics', description: 'Measures deposit increase over period.', syntax: '=(Current-Base)/Base', example: '₹4042 Cr to ₹4390 Cr → (4390-4042)/4042 = 8.6 % growth.' },
  { name: 'Advance Growth %', category: 'Bank Branch Performance Metrics', description: 'Measures loan book growth.', syntax: '=(Current-Base)/Base', example: '₹1830 Cr → ₹1985 Cr = 8.5 % increase.' },
  { name: 'CD Ratio', category: 'Bank Branch Performance Metrics', description: 'Credit to Deposit Ratio.', syntax: '=Advances/Deposits', example: '₹1985 Cr ÷ ₹4390 Cr = 45 %.' },
  { name: 'CASA %', category: 'Bank Branch Performance Metrics', description: 'Proportion of Current and Savings Accounts.', syntax: '=(Current+Savings)/Total Deposit', example: '₹362 + ₹3000 / ₹4390 = 77 % CASA.' },
  { name: 'Profit per Employee', category: 'Bank Branch Performance Metrics', description: 'Average profit generated per employee.', syntax: '=Total Profit/No of Staff', example: '₹9 Cr / 90 = ₹10 L per employee.' },
  
  { name: 'CUBEVALUE', category: 'Power Query / Data Model Helpers', description: 'Returns a measure or value from the cube/model.', syntax: '=CUBEVALUE("ThisWorkbookDataModel","[Measures].[Total Sales]")', example: 'Pull your Total Sales measure directly into a worksheet cell.' },
  { name: 'CUBEMEMBER', category: 'Power Query / Data Model Helpers', description: 'Returns a specific member from a dimension.', syntax: '=CUBEMEMBER("ThisWorkbookDataModel","[Product].[Loans]")', example: 'Reference the "Loans" product member for slicer-like reports.' },
  { name: 'CUBERANKEDMEMBER', category: 'Power Query / Data Model Helpers', description: 'Returns the nth member in a set.', syntax: '=CUBERANKEDMEMBER("ThisWorkbookDataModel",A1,1)', example: 'If A1 holds a set of branches, return the top-ranked branch.' },
  { name: 'GETPIVOTDATA (Model)', category: 'Power Query / Data Model Helpers', description: 'Fetches a value from a PivotTable linked to the Data Model.', syntax: '=GETPIVOTDATA("Sales",$A$3,"Region","East")', example: 'Pull East region Sales from a Pivot connected to the model.' },


  { name: 'DATA VALIDATION (LIST)', category: 'Data Validation & List Automation', description: 'Creates a dropdown list for controlled data entry.', syntax: '=Data Validation → List → =$A$2:$A$10', example: 'Allow only branch names (Kolaghat, Tamluk, Haldia) in a dropdown.' },
  { name: 'ISNUMBER+SEARCH', category: 'Data Validation & List Automation', description: 'Checks if a cell contains a keyword.', syntax: '=ISNUMBER(SEARCH("loan",B2))', example: 'Flags TRUE when “loan” appears in remarks (e.g., “Gold Loan Closed”).' },
  { name: 'COUNTIF VALIDATION', category: 'Data Validation & List Automation', description: 'Prevents duplicates by checking existing values.', syntax: '=COUNTIF($A$2:$A$100,A2)=1', example: 'Stops duplicate Customer IDs in data entry sheets.' },
  { name: 'LEN', category: 'Data Validation & List Automation', description: 'Counts characters in a cell.', syntax: '=LEN(A2)', example: 'Ensure PAN number has exactly 10 characters.' },
  { name: 'LEFT+RIGHT+MID', category: 'Data Validation & List Automation', description: 'Extracts text portions for validation.', syntax: '=LEFT(A2,5)', example: 'Use to check first 5 digits of an account number.' },


  { name: 'STOCK VALUE', category: 'Inventory & Procurement Management', description: 'Calculates total inventory value.', syntax: '=Quantity*Rate', example: '100 chairs × ₹800 each = ₹80,000.' },
  { name: 'REORDER LEVEL', category: 'Inventory & Procurement Management', description: 'Level at which new order is needed.', syntax: '=Average Daily Use*Lead Time', example: '50 units/day × 7 days = 350 units reorder point.' },
  { name: 'SAFETY STOCK', category: 'Inventory & Procurement Management', description: 'Extra stock to avoid stock-out.', syntax: '=(Max Demand-Average Demand)*Lead Time', example: '60-50 × 5 days = 50 units safety buffer.' },
  { name: 'AVERAGE AGE OF STOCK', category: 'Inventory & Procurement Management', description: 'Shows how long items stay in inventory.', syntax: '=(Avg Inventory/COGS)*365', example: '₹3 L avg stock / ₹12 L COGS × 365 = 91 days.' },
  { name: 'INVENTORY TURNOVER', category: 'Inventory & Procurement Management', description: 'How many times stock is sold in a year.', syntax: '=COGS/Avg Inventory', example: '₹12 L / ₹3 L = 4 times per year.' },


  { name: 'XLOOKUP', category: 'Lookup & Reference', description: 'Modern lookup both directions with defaults.', syntax: '=XLOOKUP(A2,CustomerID,CustomerName,"Not Found")', example: 'Fetches customer name for given ID, else shows Not Found.' },

  { name: 'TEXT', category: 'Text Formatting & Number Display', description: 'Displays a number in a custom format (currency, dates, percentages, etc.).', syntax: '=TEXT(A2,"₹#,##0.00")', example: 'If A2 = 12500, result shows ₹12,500.00 on your invoice.' },
  { name: 'CONCAT + TEXT', category: 'Text Formatting & Number Display', description: 'Combines text with a formatted number.', syntax: '="Target Achieved: "&TEXT(B2,"0.00%")', example: 'If B2 = 0.875, result is "Target Achieved: 87.50%".' },
  { name: 'ROMAN', category: 'Text Formatting & Number Display', description: 'Converts Arabic numbers to Roman numerals.', syntax: '=ROMAN(2025)', example: 'Useful for report versioning: 2025 → "MMXXV".' },
  { name: 'DOLLAR', category: 'Text Formatting & Number Display', description: 'Formats a number as US currency text.', syntax: '=DOLLAR(A2,2)', example: 'If A2 = 1540.5, result = "$1,540.50" for a USD export.' },
  { name: 'BASE', category: 'Text Formatting & Number Display', description: 'Converts a decimal number to another base (2–36).', syntax: '=BASE(31,2)', example: 'For internal IDs, 31 in base 2 → "11111".' },


  { name: 'Disbursement Rate', category: 'MSME Loan Monitoring', description: 'Disbursed amount as % of sanction.', syntax: '=Disbursed/Sanctioned', example: '₹80 L of ₹1 Cr → 80 % disbursement.' },
  { name: 'NPA Ratio', category: 'MSME Loan Monitoring', description: 'Measures non-performing assets.', syntax: '=NPA/Total Loan', example: '₹3 Cr NPA / ₹120 Cr Advances = 2.5 %.' },
  { name: 'Recovery %', category: 'MSME Loan Monitoring', description: 'Collections as % of overdue amount.', syntax: '=Recovered/Overdue', example: '₹50 L collected of ₹80 L → 62.5 %.' },
  { name: 'Average Ticket Size', category: 'MSME Loan Monitoring', description: 'Average loan value per account.', syntax: '=Total Loan/No of Accounts', example: '₹120 Cr / 800 = ₹15 L average size.' },
  { name: 'Turnaround Time', category: 'MSME Loan Monitoring', description: 'Average processing time for loans.', syntax: '=AVERAGE(Days)', example: 'Average = 5 days from application to disbursement.' },


  { name: 'Crop Loan Share %', category: 'Agriculture Credit Analysis', description: 'Proportion of crop loan in total agri advances.', syntax: '=Crop Loan/Total Agri Loan', example: '₹30 Cr / ₹50 Cr = 60 %.' },
  { name: 'KCC Utilization', category: 'Agriculture Credit Analysis', description: 'Average usage of Kisan Credit Card limit.', syntax: '=Drawn Amount/Sanction Limit', example: '₹3 L drawn of ₹5 L limit = 60 % utilized.' },
  { name: 'Agri NPA %', category: 'Agriculture Credit Analysis', description: 'Ratio of non-performing agri loans.', syntax: '=Agri NPA/Total Agri Advance', example: '₹1 Cr NPA / ₹40 Cr advances = 2.5 %.' },
  { name: 'Agri Growth YoY', category: 'Agriculture Credit Analysis', description: 'Year-on-year growth in agri portfolio.', syntax: '=(Current-Previous)/Previous', example: '₹40 Cr → ₹44 Cr = 10 % growth.' },
  { name: 'SHG Loan Linkage', category: 'Agriculture Credit Analysis', description: 'Number of Self-Help Groups linked to credit.', syntax: '=Linked SHG/Total SHG', example: '450 linked of 500 → 90 % credit linkage.' },


  { name: 'Deposit Growth', category: 'Deposit Mobilization & Analysis', description: 'Growth in deposits over period.', syntax: '=(Current-Base)/Base', example: '₹4042 → ₹4390 Cr = 8.6 % growth.' },
  { name: 'Average Deposit per Account', category: 'Deposit Mobilization & Analysis', description: 'Average balance held per account.', syntax: '=Total Deposit/No of Accounts', example: '₹4390 Cr / 4.2 L = ₹1.04 L average.' },
  { name: 'CASA Share %', category: 'Deposit Mobilization & Analysis', description: 'Share of Current + Savings deposits.', syntax: '=(CA+SA)/Total Deposit', example: '₹3362 / ₹4390 = 76.6 %.' },
  { name: 'Term Deposit Share %', category: 'Deposit Mobilization & Analysis', description: 'Term deposits as % of total deposit.', syntax: '=Term/Total Deposit', example: '₹1028 / ₹4390 = 23.4 %.' },
  { name: 'Average Deposit Growth %', category: 'Deposit Mobilization & Analysis', description: 'Average annual growth rate.', syntax: '=(Current/Base)^(1/Years)-1', example: '8.6 % growth in one year → 8.6 % CAGR.' },


  { name: 'Weighted Score', category: 'Branch Ranking & Scorecard', description: 'Calculates weighted performance marks.', syntax: '=SUMPRODUCT(Marks,Weights)/SUM(Weights)', example: 'Combine Deposit, Advance, CASA weights to rank branches.' },
  { name: 'Rank', category: 'Branch Ranking & Scorecard', description: 'Finds rank among branches.', syntax: '=RANK.EQ(Total,Totals)', example: 'Assign rank based on total score out of 27 branches.' },
  { name: 'Achievement %', category: 'Branch Ranking & Scorecard', description: 'Performance vs target.', syntax: '=Actual/Target', example: '₹950 Cr of ₹1000 Cr → 95 % achieved.' },
  { name: 'Gap to Target', category: 'Branch Ranking & Scorecard', description: 'Remaining gap to reach goal.', syntax: '=Target-Actual', example: '₹1000 Cr - ₹950 Cr = ₹50 Cr gap.' },
  { name: 'Overall Rank Change', category: 'Branch Ranking & Scorecard', description: 'Compares rank change period to period.', syntax: '=Previous Rank-Current Rank', example: 'Rank improved from 10 → 5 = +5 gain.' },


  { name: 'Business per Employee', category: 'Operational Efficiency Metrics', description: 'Average business handled per staff.', syntax: '=(Deposits+Advances)/No of Staff', example: '₹6200 Cr / 90 = ₹68.9 Cr per employee.' },
  { name: 'Profit per Branch', category: 'Operational Efficiency Metrics', description: 'Average profit generated per branch.', syntax: '=Total Profit/No of Branches', example: '₹120 Cr profit / 79 branches = ₹1.52 Cr per branch.' },
  { name: 'Cost to Income Ratio', category: 'Operational Efficiency Metrics', description: 'Expense efficiency measure.', syntax: '=Operating Expense/Operating Income', example: '₹40 Cr / ₹100 Cr = 40 % ratio.' },
  { name: 'Staff Productivity Index', category: 'Operational Efficiency Metrics', description: 'Composite productivity score.', syntax: '=Business per Employee/Avg Business Benchmark', example: '₹68.9 Cr vs ₹60 Cr benchmark → 115 % index.' },
  { name: 'Transaction per Employee', category: 'Operational Efficiency Metrics', description: 'Total transactions divided by staff.', syntax: '=Total Txn/No of Employees', example: '18,000 transactions / 90 = 200 per staff.' },


  { name: 'Net Profit', category: 'Profit & Loss Computation', description: 'Income minus expenses.', syntax: '=Total Income-Total Expense', example: '₹120 Cr – ₹85 Cr = ₹35 Cr profit.' },
  { name: 'Gross Profit Margin', category: 'Profit & Loss Computation', description: 'Profit as a percentage of revenue.', syntax: '=(Sales-COGS)/Sales', example: '₹12 L – ₹9 L = 25 % margin.' },
  { name: 'Operating Profit', category: 'Profit & Loss Computation', description: 'Earnings before interest and taxes.', syntax: '=EBIT', example: 'Shows core operating profit before finance costs.' },
  { name: 'Net Margin %', category: 'Profit & Loss Computation', description: 'Net profit as a share of income.', syntax: '=Net Profit/Total Income', example: '₹35 Cr / ₹120 Cr = 29.1 %.' },
  { name: 'Break-Even Sales', category: 'Profit & Loss Computation', description: 'Sales needed to cover fixed costs.', syntax: '=Fixed Cost/(Selling Price-Variable Cost)', example: '₹40 L / (₹500-₹350) = 26,667 units.' },


  { name: 'Budget Variance', category: 'Expense & Budget Monitoring', description: 'Difference between budgeted and actual spend.', syntax: '=Budget-Actual', example: 'Budget ₹5 L, spent ₹4.6 L → ₹0.4 L saving.' },
  { name: 'Variance %', category: 'Expense & Budget Monitoring', description: 'Percentage of budget used.', syntax: '=(Actual/Budget)-1', example: '₹4.6 L / ₹5 L - 1 = -8 % under-budget.' },
  { name: 'Cumulative Spend', category: 'Expense & Budget Monitoring', description: 'Total spent to date.', syntax: '=SUM(Actual Range)', example: 'Sum of monthly expenses Jan–Sep → ₹42 L.' },
  { name: 'Spend per Branch', category: 'Expense & Budget Monitoring', description: 'Average expense across branches.', syntax: '=Total Expense/No of Branches', example: '₹42 L / 79 branches = ₹53,000 per branch.' },
  { name: 'Projected Year-End Spend', category: 'Expense & Budget Monitoring', description: 'Forecast spend using current trend.', syntax: '=Current Spend*(12/Months Completed)', example: '₹42 L in 9 months → projected ₹56 L annual spend.' },

  { name: 'IPMT (Range)', category: 'Banking & Loan Analysis', description: 'Returns the interest amount for each period of a loan.', syntax: '=IPMT(rate/12, ROW(A1:A12), nper, -principal)', example: 'For a ₹10,00,000 loan @ 9% for 10 years (120 months), fill a column with =IPMT(9%/12,ROW(A1:A12),120,-1000000) to list interest for the first 12 EMIs.' },
  { name: 'PPMT (Range)', category: 'Banking & Loan Analysis', description: 'Returns the principal amount for each period of a loan.', syntax: '=PPMT(rate/12, ROW(A1:A12), nper, -principal)', example: 'Using the same loan, =PPMT(9%/12,ROW(A1:A12),120,-1000000) lists principal portion for each of the first 12 EMIs.' },
  { name: 'CUMIPMT', category: 'Banking & Loan Analysis', description: 'Total interest paid between two periods.', syntax: '=CUMIPMT(rate/12, nper, principal, start_period, end_period, 0)', example: 'Total interest paid in year 1 on a ₹10,00,000 loan @ 9%/yr: =CUMIPMT(9%/12,120,1000000,1,12,0).' },
  { name: 'CUMPRINC', category: 'Banking & Loan Analysis', description: 'Total principal repaid between two periods.', syntax: '=CUMPRINC(rate/12, nper, principal, start_period, end_period, 0)', example: 'Total principal repaid in year 1: =CUMPRINC(9%/12,120,1000000,1,12,0).' },
  { name: 'SYD', category: 'Banking & Loan Analysis', description: 'Sum-of-years’ digits depreciation.', syntax: '=SYD(cost, salvage, life, period)', example: 'Asset ₹5,00,000, salvage ₹50,000, 5-year life — first-year depreciation: =SYD(500000,50000,5,1).' },
  { name: 'DDB', category: 'Banking & Loan Analysis', description: 'Double-declining balance depreciation.', syntax: '=DDB(cost, salvage, life, period, [factor])', example: 'Same asset using double-declining factor 2: =DDB(500000,50000,5,1,2).' },
  { name: 'COUNTIF (Duplicate Check)', category: 'Data Analysis Helper', description: 'Counts how many times a value appears.', syntax: '=COUNTIF(A:A, A2)', example: 'In a Customer ID column, highlight duplicates where =COUNTIF(A:A,A2)>1.' },
  { name: 'UNIQUE + COUNTIF (Frequency)', category: 'Data Analysis Helper', description: 'Generates frequency table of values.', syntax: '=COUNTIF(A:A, UNIQUE(A:A))', example: 'List each product once and show how many times it was sold.' },
  { name: 'FILTER + SEARCH (Text Match)', category: 'Data Analysis Helper', description: 'Filters rows where a keyword is found.', syntax: '=FILTER(A2:D100, ISNUMBER(SEARCH("overdue", D2:D100)))', example: 'Extract only those tickets whose remarks contain “overdue”.' },
  { name: 'ISFORMULA', category: 'Data Analysis Helper', description: 'Checks if a cell contains a formula.', syntax: '=ISFORMULA(E5)', example: 'Audit KPI cells to ensure they are formula-driven, not typed values.' },
  { name: 'FORMULATEXT + IFERROR', category: 'Data Analysis Helper', description: 'Displays formula or friendly text if none.', syntax: '=IFERROR(FORMULATEXT(E5),"No formula")', example: 'Show “No formula” if the KPI cell was typed by mistake.' },
  { name: 'SORT + TRANSPOSE', category: 'Data Analysis Helper', description: 'Reshapes and sorts a list.', syntax: '=SORT(TRANSPOSE(A2:A20))', example: 'Turn a vertical list of names into a sorted horizontal header row.' },
  { name: 'CONVERT', category: 'Engineering, Scientific & Math', description: 'Converts a number from one unit to another.', syntax: '=CONVERT(value, "from_unit", "to_unit")', example: 'Convert 10 miles to kilometers: =CONVERT(10,"mi","km") → 16.093.' },
  { name: 'BESSELI', category: 'Engineering, Scientific & Math', description: 'Modified Bessel function of the first kind.', syntax: '=BESSELI(x, n)', example: 'Specialized math; for x=1.5, n=2: =BESSELI(1.5,2).' },
  { name: 'COMPLEX', category: 'Engineering, Scientific & Math', description: 'Creates a complex number as text.', syntax: '=COMPLEX(real, imaginary, ["i"|"j"])', example: '=COMPLEX(3,4,"i") returns "3+4i".' },
  { name: 'IMABS', category: 'Engineering, Scientific & Math', description: 'Absolute value (modulus) of a complex number.', syntax: '=IMABS(complex_text)', example: '=IMABS("3+4i") → 5.' },
  { name: 'DELTA', category: 'Engineering, Scientific & Math', description: 'Returns 1 if numbers are equal; otherwise 0.', syntax: '=DELTA(A2,B2)', example: 'Check if two sensor readings match exactly.' },
  { name: 'ERF', category: 'Engineering, Scientific & Math', description: 'Error function for probability/statistics.', syntax: '=ERF(x)', example: 'Used in statistical process control; e.g., =ERF(1.5).' },
  { name: 'Table.Distinct (PQ)', category: 'Power Query / Data Model', description: 'Removes duplicate rows in Power Query.', syntax: 'Table.Distinct(Source)', example: 'In PQ, remove duplicate Customer IDs before loading to Excel.' },
  { name: 'Table.Sort (PQ)', category: 'Power Query / Data Model', description: 'Sorts a table by one or more columns.', syntax: 'Table.Sort(Source,{{"Amount",Order.Descending}})', example: 'Sort transactions by Amount descending during ETL.' },
  { name: 'Table.Group (PQ)', category: 'Power Query / Data Model', description: 'Groups rows and aggregates.', syntax: 'Table.Group(Source, {"Region"}, {{"Total", each List.Sum([Amount]), type number}})', example: 'Total Amount by Region before building a Pivot.' },
  { name: 'List.Accumulate (PQ)', category: 'Power Query / Data Model', description: 'Accumulates values in a list (reduce).', syntax: 'List.Accumulate(List, seed, (state, current) => ...)', example: 'Create running totals in PQ for daily sales.' },
  { name: 'Record.Field (PQ)', category: 'Power Query / Data Model', description: 'Gets a field value from a record.', syntax: 'Record.Field(Record, "FieldName")', example: 'Pick a single attribute (like Rate) from a JSON record.' },
  { name: 'Number.Round (PQ)', category: 'Power Query / Data Model', description: 'Rounds a number to specified digits.', syntax: 'Number.Round(value, digits)', example: 'Round imported tax amounts to 2 decimals during data cleaning.' },
  { name: 'Dynamic Chart Title', category: 'Dashboard Formula Tricks', description: 'Shows current period in chart title.', syntax: '="Sales Trend – "&TEXT(TODAY(),"mmm yyyy")', example: 'On Oct 2025 it shows “Sales Trend – Oct 2025”.' },
  { name: 'Conditional Label', category: 'Dashboard Formula Tricks', description: 'Shows up/down arrow by target.', syntax: '=IF(A2>=Target,"▲ Above","▼ Below")', example: 'If October sales meet the target, title shows “▲ Above”.' },
  { name: 'In-Cell Progress Bar', category: 'Dashboard Formula Tricks', description: 'Blocks representing progress.', syntax: '=REPT("█", ROUND(A2/100*10,0))', example: 'For 70% completion → "███████".' },
  { name: 'KPI Status', category: 'Dashboard Formula Tricks', description: 'Traffic-light status using IFS.', syntax: '=IFS(A2>=Target,"Good",A2=Target,"On Target",A2<Target,"Low")', example: 'Marks branch as Good/On Target/Low.' },
  { name: 'NA() for Chart Gap', category: 'Dashboard Formula Tricks', description: 'Force breaks in line charts.', syntax: '=IF(A2=0,NA(),A2)', example: 'Missing month data shows a visual gap instead of a zero.' },
  { name: 'INFO("memavail")', category: 'Rarely Known & Handy', description: 'Returns available memory info (legacy behavior).', syntax: '=INFO("memavail")', example: 'Quick diagnostic while handling very large workbooks.' },
  { name: 'SORT(RANDARRAY())', category: 'Rarely Known & Handy', description: 'Generates a random sorted list.', syntax: '=SORT(RANDARRAY(10,1,1,100,TRUE))', example: 'Create a random order of 10 employee IDs for audits.' },
  { name: 'LEN(FORMULATEXT())', category: 'Rarely Known & Handy', description: 'Counts characters in a formula.', syntax: '=LEN(FORMULATEXT(A1))', example: 'Find very long/complex formulas to refactor.' },
  { name: 'CONCAT + CHAR(10)', category: 'Rarely Known & Handy', description: 'Joins text onto multiple lines within a cell.', syntax: '=CONCAT(A2,CHAR(10),B2)', example: 'Show customer name and address on two lines in one cell.' },
  { name: 'SEQUENCE + TODAY', category: 'Rarely Known & Handy', description: 'Creates a list of consecutive dates.', syntax: '=SEQUENCE(7,1,TODAY(),1)', example: 'Generate dates for the next 7 days for daily tasks.' },
  { name: 'IMAGE', category: 'Rarely Known & Handy', description: 'Displays an image from a URL in a cell (Excel 365).', syntax: '=IMAGE("https://example.com/logo.png")', example: 'Show company logo in a dashboard header.' },
  { name: 'HYPERLINK to Slide', category: 'Excel → PowerPoint', description: 'Jump to a specific slide (with named anchors).', syntax: '=HYPERLINK("#Slide1","Go to Intro")', example: 'Click cell to jump to Slide 1 in a linked deck.' },
  { name: 'TEXTJOIN Notes', category: 'Excel → PowerPoint', description: 'Combines bullet items into a single cell for export.', syntax: '=TEXTJOIN(CHAR(10),TRUE,A2:A6)', example: 'Create slide notes with each bullet on a new line.' },
  { name: 'Dynamic Chart Caption', category: 'Excel → PowerPoint', description: 'Automatically updates chart caption with month.', syntax: '="Sales "&TEXT(TODAY(),"mmm yy")', example: 'Shows “Sales Oct 25” under a chart you export.' },
  { name: 'Publish Add-in (Tip)', category: 'Excel → PowerPoint', description: 'Use “Export to PowerPoint/Designer” to push charts and tables.', syntax: 'N/A (Add-in/Feature)', example: 'Export KPI table and sparkline chart to a meeting deck.' },
  { name: 'WEBSERVICE (API call)', category: 'Excel + AI Integration', description: 'Calls a web API endpoint (legacy; XML/JSON text).', syntax: '=WEBSERVICE("https://api.example.com/endpoint?key=XXXX")', example: 'Fetch a short analysis string that you then parse with FILTERXML/TEXT functions.' },
  { name: 'FILTERXML (parse)', category: 'Excel + AI Integration', description: 'Parses XML (or XML-shaped JSON) from a text cell.', syntax: '=FILTERXML(A1,"//item")', example: 'Extract the “summary” nodes returned by an AI microservice into cells.' },
  { name: 'TEXTJOIN → JSON', category: 'Excel + AI Integration', description: 'Builds JSON payloads from cells.', syntax: '=TEXTJOIN("",TRUE,"{""query"":""",A2,""",""id"":",B2,"}")', example: 'Turn cell values into a valid JSON string for scripts.' },
  { name: 'ENCODEURL', category: 'Excel + AI Integration', description: 'URL-encodes text for GET requests.', syntax: '=ENCODEURL(A2)', example: 'Encode “credit risk score” to pass safely in an API query string.' },
  { name: 'IMAGE (from AI output)', category: 'Excel + AI Integration', description: 'Embeds a returned image URL from an AI service.', syntax: '=IMAGE(A2)', example: 'Show a generated chart image or product mock inside a cell.' },
  { name: 'ANALYZE', category: 'Excel 365 AI (Preview)', description: 'Summarizes a selected table and highlights drivers (preview feature).', syntax: '=ANALYZE(A1:D200)', example: 'Run on a sales table to get plain-English trends and outliers.' },
  { name: 'SUMMARIZE', category: 'Excel 365 AI (Preview)', description: 'Creates a short narrative summary of a range.', syntax: '=SUMMARIZE(A1:D200)', example: 'Explain monthly spending vs budget in a paragraph.' },
  { name: 'EXPLAIN', category: 'Excel 365 AI (Preview)', description: 'Explains how a complex formula works.', syntax: '=EXPLAIN(F5)', example: 'Paste in an INDEX/MATCH/LAMBDA combo and get a human explanation.' },
  { name: 'GENERATE', category: 'Excel 365 AI (Preview)', description: 'Generates a formula from a natural-language prompt.', syntax: '=GENERATE("Return top 5 branches by deposits")', example: 'Returns a valid formula using SORT/TOP logic to show top branches.' },
  { name: 'INSIGHT', category: 'Excel 365 AI (Preview)', description: 'Detects patterns and anomalies.', syntax: '=INSIGHT(A1:D200)', example: 'Highlights a sudden spike in cash deposits in September.' },
  { name: 'SEARCH (Keyword Match)', category: 'Index & Search Helpers', description: 'Finds the position of a keyword inside text.', syntax: '=SEARCH("loan", A2)', example: 'Returns a number if “loan” appears anywhere in A2; use with ISNUMBER to filter help topics.' },
  { name: 'FILTER (Search Results)', category: 'Index & Search Helpers', description: 'Returns rows where a search condition is true.', syntax: '=FILTER(A2:D500, ISNUMBER(SEARCH($H$1, B2:B500)))', example: 'Type a keyword in H1 and instantly list matching formulas or topics.' },
  { name: 'SORTBY (Relevance/Date)', category: 'Index & Search Helpers', description: 'Sorts results by another column (e.g., relevance or date).', syntax: '=SORTBY(ResultRange, ScoreRange, -1)', example: 'Sort help results by highest relevance score or latest update.' },
  { name: 'TEXTJOIN (Snippet Build)', category: 'Index & Search Helpers', description: 'Builds preview text from multiple fields.', syntax: '=TEXTJOIN(" – ", TRUE, Name, Category, Syntax)', example: 'Show “VLOOKUP – Lookup – =VLOOKUP(...)” in a single preview cell.' },
  { name: 'UNIQUE (Category List)', category: 'Index & Search Helpers', description: 'Generates a unique list of categories for filters.', syntax: '=UNIQUE(CategoryRange)', example: 'Create a filter dropdown of distinct categories: Lookup, Text, Financial, etc.' },

  { name:'RIGHT', category:'Text', description:'Extracts the last characters from a text string.', syntax:'=RIGHT(text,[num_chars])', example:'Account "1234567890" → =RIGHT(A2,4) returns "7890".' },
  { name:'MID', category:'Text', description:'Returns a specific number of characters from the middle of a text string.', syntax:'=MID(text,start_num,num_chars)', example:'Email "amit.dutta@pnb.co.in" → =MID(A2,6,5) returns "dutta".' },
  { name:'UNICODE', category:'Text', description:'Returns the Unicode number of the first character of text.', syntax:'=UNICODE(text)', example:'=UNICODE("₹") returns 8377.' },
  { name:'UNICHAR', category:'Text', description:'Returns the character referred to by the Unicode number.', syntax:'=UNICHAR(number)', example:'=UNICHAR(8377) returns "₹".' },

  { name:'HOUR', category:'Date & Time', description:'Extracts the hour from a time value.', syntax:'=HOUR(serial_number)', example:'If A2=10:45 AM, =HOUR(A2) returns 10.' },
  { name:'MINUTE', category:'Date & Time', description:'Extracts the minute from a time value.', syntax:'=MINUTE(serial_number)', example:'If A2=9:25 PM, =MINUTE(A2) returns 25.' },
  { name:'SECOND', category:'Date & Time', description:'Extracts the second from a time value.', syntax:'=SECOND(serial_number)', example:'If A2=12:30:45, =SECOND(A2) returns 45.' },

  { name:'MODE.SNGL', category:'Statistical', description:'Returns the most frequently occurring value in a set.', syntax:'=MODE.SNGL(number1,[number2],...)', example:'Test scores in A2:A20 → =MODE.SNGL(A2:A20) gives the most common score.' },
  { name:'MODE.MULT', category:'Statistical', description:'Returns a vertical array of the most frequent values.', syntax:'=MODE.MULT(number1,[number2],...)', example:'Sales amounts in A2:A20 with two peaks → =MODE.MULT(A2:A20) spills both values.' },
  { name:'QUARTILE.INC', category:'Statistical', description:'Returns the quartile of a dataset including endpoints.', syntax:'=QUARTILE.INC(array,quart)', example:'Deposits B2:B50 → =QUARTILE.INC(B2:B50,1) gives Q1.' },
  { name:'PERCENTRANK.INC', category:'Statistical', description:'Returns the rank of a value as a percentage (inclusive).', syntax:'=PERCENTRANK.INC(array,x,[significance])', example:'₹80,00,000 vs deposits B2:B50 → =PERCENTRANK.INC(B2:B50,8000000).' },
  { name:'CHISQ.TEST', category:'Statistical', description:'Returns the test for independence (chi-square).', syntax:'=CHISQ.TEST(actual_range,expected_range)', example:'Compare observed vs expected satisfaction counts → =CHISQ.TEST(B2:C6,D2:E6).' },
  { name:'CONFIDENCE.NORM', category:'Statistical', description:'Confidence interval for a population mean (normal).', syntax:'=CONFIDENCE.NORM(alpha,standard_dev,size)', example:'95% CI, σ=4, n=100 → =CONFIDENCE.NORM(0.05,4,100) returns 0.784.' },
  { name:'CONFIDENCE.T', category:'Statistical', description:'Confidence interval using Student’s t-distribution.', syntax:'=CONFIDENCE.T(alpha,standard_dev,size)', example:'Small sample (n=20), σ=3 → margin via =CONFIDENCE.T(0.05,3,20).' },

  { name:'ACCRINTM', category:'Financial', description:'Returns accrued interest for a security that pays at maturity.', syntax:'=ACCRINTM(issue,settlement,rate,par,[basis])', example:'Compute total interest earned between issue and maturity dates for a zero-coupon bond.' },
  { name:'ODDFYIELD', category:'Financial', description:'Returns yield of a security with an odd first coupon period.', syntax:'=ODDFYIELD(settlement,maturity,issue,first_coupon,rate,pr,redemption,frequency,[basis])', example:'Used for new bond issues with irregular first period to get comparable annualized yield.' },

  { name:'IMSUB', category:'Engineering & Complex', description:'Subtracts two complex numbers.', syntax:'=IMSUB(inumber1,inumber2)', example:'"4+3i" − "2+2i" → =IMSUB("4+3i","2+2i") returns "2+1i".' },
  { name:'IMCONJUGATE', category:'Engineering & Complex', description:'Returns the complex conjugate of a complex number.', syntax:'=IMCONJUGATE(inumber)', example:'=IMCONJUGATE("5+3i") returns "5-3i".' },
  { name:'IMREAL', category:'Engineering & Complex', description:'Returns the real coefficient of a complex number.', syntax:'=IMREAL(inumber)', example:'=IMREAL("4+6i") returns 4.' },
  { name:'IMAGINARY', category:'Engineering & Complex', description:'Returns the imaginary coefficient of a complex number.', syntax:'=IMAGINARY(inumber)', example:'=IMAGINARY("4+6i") returns 6.' },

  { name:'BIN2HEX', category:'Conversions & Bases', description:'Converts a binary number to hexadecimal.', syntax:'=BIN2HEX(number,[places])', example:'=BIN2HEX("1111") returns "F".' },
  { name:'HEX2BIN', category:'Conversions & Bases', description:'Converts a hexadecimal number to binary.', syntax:'=HEX2BIN(number,[places])', example:'=HEX2BIN("FF") returns "11111111".' },
  { name:'BIN2OCT', category:'Conversions & Bases', description:'Converts a binary number to octal.', syntax:'=BIN2OCT(number,[places])', example:'=BIN2OCT("1010") returns "12".' },
  { name:'OCT2BIN', category:'Conversions & Bases', description:'Converts an octal number to binary.', syntax:'=OCT2BIN(number,[places])', example:'=OCT2BIN("12") returns "1010".' },
  { name:'OCT2HEX', category:'Conversions & Bases', description:'Converts an octal number to hexadecimal.', syntax:'=OCT2HEX(number,[places])', example:'=OCT2HEX("77") returns "3F".' },
  { name:'HEX2OCT', category:'Conversions & Bases', description:'Converts a hexadecimal number to octal.', syntax:'=HEX2OCT(number,[places])', example:'=HEX2OCT("FF") returns "377".' },

  { name:'PRODUCT', category:'Math & Trig', description:'Multiplies all the numbers given as arguments and returns the product.', syntax:'=PRODUCT(number1,[number2],...)', example:'Quantities in A2:A5 → =PRODUCT(A2:A5) multiplies all values together.' },

  { name:'RADIANS', category:'Measurement & Unit Conversion', description:'Converts degrees to radians.', syntax:'=RADIANS(angle)', example:'=RADIANS(180) returns 3.1416.' },
  { name:'DEGREES', category:'Measurement & Unit Conversion', description:'Converts radians to degrees.', syntax:'=DEGREES(angle)', example:'=DEGREES(PI()) returns 180.' },
  { name:'SIN', category:'Measurement & Unit Conversion', description:'Returns the sine of an angle in radians.', syntax:'=SIN(angle)', example:'=SIN(RADIANS(30)) returns 0.5.' },
  { name:'COS', category:'Measurement & Unit Conversion', description:'Returns the cosine of an angle in radians.', syntax:'=COS(angle)', example:'=COS(RADIANS(60)) returns 0.5.' },
  { name:'TAN', category:'Measurement & Unit Conversion', description:'Returns the tangent of an angle in radians.', syntax:'=TAN(angle)', example:'=TAN(RADIANS(45)) returns 1.' },

  { name:'CONVERT_C_TO_F', category:'Temperature & Physical Constants', description:'Converts temperature from Celsius to Fahrenheit.', syntax:'=(Celsius*9/5)+32', example:'30°C → =(30*9/5)+32 returns 86°F.' },
  { name:'CONVERT_F_TO_C', category:'Temperature & Physical Constants', description:'Converts temperature from Fahrenheit to Celsius.', syntax:'=(Fahrenheit-32)*5/9', example:'98°F → =(98-32)*5/9 returns 36.7°C.' },
  { name:'CONVERT_C_TO_K', category:'Temperature & Physical Constants', description:'Converts Celsius to Kelvin.', syntax:'=Celsius+273.15', example:'25°C → =25+273.15 returns 298.15K.' },
  { name:'CONVERT_K_TO_C', category:'Temperature & Physical Constants', description:'Converts Kelvin to Celsius.', syntax:'=Kelvin-273.15', example:'310K → =310-273.15 returns 36.85°C.' },
  { name:'ABSOLUTE_ZERO', category:'Temperature & Physical Constants', description:'Reference value for absolute zero temperature.', syntax:'=0', example:'0K equals −273.15°C (theoretical minimum temperature).' },
  { name:'G_CONSTANT', category:'Temperature & Physical Constants', description:'Gravitational constant (for physics formulas).', syntax:'=6.674E-11', example:'Used in F = G*(m1*m2)/r^2.' },
  { name:'SPEED_OF_LIGHT', category:'Temperature & Physical Constants', description:'Speed of light in m/s.', syntax:'=2.998E8', example:'Used in E = m*c^2, c = 2.998×10^8 m/s.' },
  { name:'PLANCK_CONSTANT', category:'Temperature & Physical Constants', description:'Planck’s constant in J·s.', syntax:'=6.626E-34', example:'Photon energy E = h*f.' },
  { name:'AVOGADRO_NUMBER', category:'Temperature & Physical Constants', description:'Avogadro’s number (particles per mole).', syntax:'=6.022E23', example:'Number of molecules in 1 mole.' },
  { name:'BOLTZMANN_CONSTANT', category:'Temperature & Physical Constants', description:'Links temperature to energy (J/K).', syntax:'=1.381E-23', example:'Used in gas and thermal equations.' },

  { name:'ELECTRICITY_BILL', category:'Household & Utility', description:'Monthly electricity bill from units, per-unit rate, and fixed charge.', syntax:'=Units*kWh_Rate + Fixed_Charge', example:'300 units @ ₹7 with ₹120 fixed → =300*7+120 = ₹2,220.' },
  { name:'APPLIANCE_ENERGY_COST', category:'Household & Utility', description:'Cost to run an appliance based on wattage and hours.', syntax:'=(Watt*Hours_Per_Day*Days/1000)*kWh_Rate', example:'750W heater, 2h/day for 30 days @ ₹7 → ₹315.' },
  { name:'WATER_USAGE_PER_DAY', category:'Household & Utility', description:'Average daily water usage from monthly total.', syntax:'=Total_Monthly_Litres/Days', example:'9,000 L in 30 days → 300 L/day.' },
  { name:'WATER_BILL', category:'Household & Utility', description:'Water bill from usage and rate.', syntax:'=Usage_Litres*Rate_Per_Litre', example:'8,000 L @ ₹0.002 → ₹16.' },
  { name:'FUEL_COST_PER_KM', category:'Household & Utility', description:'Fuel cost per kilometer.', syntax:'=Fuel_Price_Per_Litre/Mileage_km_per_L', example:'₹105/L, 17 km/L → ₹6.18/km.' },
  { name:'TRIP_FUEL_COST', category:'Household & Utility', description:'Fuel cost for a trip distance.', syntax:'=Distance_km/Mileage*Fuel_Price', example:'150 km, 17 km/L, ₹105/L → ≈ ₹926.' },
  { name:'LPG_CYLINDER_DAYS', category:'Household & Utility', description:'Days a gas cylinder will last.', syntax:'=Cylinder_kg/Daily_Use_kg', example:'14.2 kg at 0.4 kg/day → 35.5 days.' },
  { name:'INTERNET_DAILY_LIMIT', category:'Household & Utility', description:'Average daily data from plan.', syntax:'=Plan_GB/Validity_Days', example:'84 GB / 28 days → 3 GB/day.' },
  { name:'RECHARGE_COST_PER_DAY', category:'Household & Utility', description:'Effective daily cost of mobile plan.', syntax:'=Plan_Cost/Validity_Days', example:'₹719 / 84 days → ₹8.56/day.' },
  { name:'MAINTENANCE_PER_MONTH', category:'Household & Utility', description:'Monthly share of annual maintenance.', syntax:'=Annual_Maintenance/12', example:'₹18,000 yearly → ₹1,500/month.' },
  { name:'GROCERY_BUDGET_VARIANCE', category:'Household & Utility', description:'Over/under spend vs budget.', syntax:'=Actual_Spend - Budget', example:'₹9,200 spent vs ₹8,500 budget → ₹700 over.' },
  { name:'WASHING_COST_PER_LOAD', category:'Household & Utility', description:'Water + electricity cost per laundry load.', syntax:'=(Water_L*Water_Rate)+(kWh_Used*kWh_Rate)', example:'60 L @ ₹0.002 + 0.8 kWh @ ₹7 → ₹5.72.' },
  { name:'SOLAR_SAVINGS_PER_MONTH', category:'Household & Utility', description:'Monthly bill savings from solar generation.', syntax:'=Solar_kWh*Grid_Rate', example:'150 kWh @ ₹7 → ₹1,050 saved.' },
  { name:'UPS_BACKUP_HOURS', category:'Household & Utility', description:'Estimated UPS backup time.', syntax:'=(Battery_Ah*Battery_V*Efficiency)/Load_Watts', example:'150Ah*12V*0.8 / 120W → 12 hours.' },

{ name:'CHOOSE', category:'Decision Support', description:'Selects a value from a list based on its position number.', syntax:'=CHOOSE(2,"Agri","Retail","MSME")', example:'Returns “Retail” as it is the 2nd option.' },
{ name:'RANDARRAY', category:'Simulation', description:'Generates an array of random numbers between two limits.', syntax:'=RANDARRAY(5,1,1,100,TRUE)', example:'Creates a list of 5 random whole numbers between 1 and 100.' },
{ name:'ExpectedValue', category:'Simulation', description:'Calculates expected value based on probabilities and outcomes.', syntax:'=SUMPRODUCT(Probabilities,Outcomes)', example:'If 60% chance to earn ₹100 and 40% to earn ₹50, result = ₹80.' },
{ name:'DATA TABLE', category:'Scenario Analysis', description:'Performs two-way what-if analysis to evaluate different outcomes.', syntax:'=DATA TABLE(RowInput,ColumnInput)', example:'Used to compare EMI for multiple interest rates and tenures.' },
{ name:'GOAL SEEK', category:'Scenario Analysis', description:'Finds the input required to achieve a target output.', syntax:'=GOAL SEEK(TargetCell,InputCell,GoalValue)', example:'Used to calculate sales needed for ₹1 Lakh profit.' },
{ name:'WEBSERVICE', category:'Web & API', description:'Fetches data directly from a web API or URL.', syntax:'=WEBSERVICE("https://api.exchangerate.host/latest")', example:'Displays live currency rates for financial dashboard.' },
{ name:'FILTERXML', category:'Web & API', description:'Extracts structured data (like XML) returned by a web service.', syntax:'=FILTERXML(A1,"//USD")', example:'Pulls USD rate from XML stored in A1.' },
{ name:'ENCODEURL', category:'Web & API', description:'Converts text into a web-safe URL format.', syntax:'=ENCODEURL(A1)', example:'If A1="Punjab National Bank", returns Punjab%20National%20Bank.' },
{ name:'ANALYZE', category:'AI & Copilot', description:'Summarizes a table and identifies data insights automatically.', syntax:'=ANALYZE(A1:D20)', example:'Highlights top-performing branches automatically (Copilot feature).' },
{ name:'EXPLAIN', category:'AI & Copilot', description:'Explains the meaning of a complex formula.', syntax:'=EXPLAIN(A1)', example:'Describes how a long nested IF formula works in simple words.' },
{ name:'SUMMARIZE', category:'AI & Copilot', description:'Creates a natural-language summary of data.', syntax:'=SUMMARIZE(A1:C10)', example:'Generates “Retail growth increased by 12% this quarter.”' },
{ name:'GENERATE', category:'AI & Copilot', description:'Produces new formulas or data patterns from context.', syntax:'=GENERATE("Calculate CAGR")', example:'Automatically inserts the correct CAGR formula.' },
{ name:'DESCRIBE', category:'AI & Copilot', description:'Provides metadata description for selected range.', syntax:'=DESCRIBE(A1:C10)', example:'Lists column headers and data types from selected range.' },
{ name:'DETERMINANT', category:'Matrix & Geometry', description:'Computes determinant of a square matrix.', syntax:'=MDETERM(A1:D4)', example:'Used in solving linear equation sets for engineering data.' },
{ name:'MINVERSE', category:'Matrix & Geometry', description:'Returns the inverse of a square matrix.', syntax:'=MINVERSE(A1:D4)', example:'Helps solve equations using matrix algebra.' },
{ name:'DOTPRODUCT', category:'Matrix & Geometry', description:'Calculates dot product between two numeric vectors.', syntax:'=SUMPRODUCT(A1:A3,B1:B3)', example:'If A={2,3,4} and B={5,6,7}, result = 56.' },
{ name:'FiscalYear', category:'Date Intelligence', description:'Computes Indian fiscal year (April–March).', syntax:'=IF(MONTH(A1)>=4,YEAR(A1),YEAR(A1)-1)', example:'If A1=15-Mar-2025, returns 2024.' },
{ name:'QuarterLabel', category:'Date Intelligence', description:'Returns fiscal quarter label (Q1-Q4).', syntax:'="Q"&INT((MONTH(A1)+2)/3)', example:'If A1=10-Jul, result = Q2.' },
{ name:'WORKDAY', category:'Date Intelligence', description:'Returns a date that is a given number of working days away.', syntax:'=WORKDAY(A1,10)', example:'If A1=1-Jan, result = 15-Jan skipping weekends.' },
{ name:'COVARIANCE.P', category:'Data Science', description:'Shows relationship between two datasets (population).', syntax:'=COVARIANCE.P(A1:A10,B1:B10)', example:'Used to compare deposit and advance growth trends.' },
{ name:'Z-SCORE', category:'Data Science', description:'Calculates how far a value is from the mean in SD units.', syntax:'=(A1-AVERAGE(A:A))/STDEV(A:A)', example:'If Z=1.5, it’s 1.5 SD above average sales.' },
{ name:'CreditDepositRatio', category:'Banking KPI', description:'Shows proportion of loans to deposits.', syntax:'=TotalAdvance/TotalDeposit', example:'If Advances ₹200 Cr, Deposits ₹250 Cr → 80%.' },
{ name:'CASARatio', category:'Banking KPI', description:'Measures share of low-cost deposits (savings + current).', syntax:'=(Savings+Current)/TotalDeposit', example:'If Savings ₹50 Cr, Current ₹20 Cr, Total ₹200 Cr → 35% CASA.' },
{ name:'WeightedGapPercent', category:'Banking KPI', description:'Measures weighted achievement against budget.', syntax:'=(Ach/Budget)*Weight', example:'If 90% achieved and weight 20, score = 18.' },
{ name:'BranchRank', category:'Banking KPI', description:'Ranks branches based on performance within circle.', syntax:'=RANK.EQ(A1,A:A)', example:'If A1=500 Cr deposit, shows branch rank among peers.' },
{ name:'IF', category:'Logical', description:'Checks a condition and returns one value for TRUE and another for FALSE.', syntax:'=IF(A1>100,"High","Low")', example:'If A1=120, result = “High”.' },

{ name:'NOT', category:'Logical', description:'Reverses the result of a logical test.', syntax:'=NOT(A1>10)', example:'If A1=5, returns TRUE.' },
{ name:'TRUE', category:'Logical', description:'Returns the logical value TRUE.', syntax:'=TRUE()', example:'Can be used to mark “valid” cells.' },
{ name:'FALSE', category:'Logical', description:'Returns the logical value FALSE.', syntax:'=FALSE()', example:'Often used in comparison formulas.' },
{ name:'SUMPRODUCT', category:'Optimization', description:'Multiplies corresponding values in arrays and sums them.', syntax:'=SUMPRODUCT(A1:A5,B1:B5)', example:'Used in Excel Solver to calculate objective function.' },
{ name:'CONVERT', category:'Measurement', description:'Converts a number from one measurement unit to another.', syntax:'=CONVERT(5,"mi","km")', example:'Converts 5 miles to 8.05 km.' },
{ name:'LCM', category:'Math', description:'Returns the least common multiple of numbers.', syntax:'=LCM(A1,B1)', example:'If A1=6 and B1=8, returns 24.' },
{ name:'GCD', category:'Math', description:'Returns the greatest common divisor of numbers.', syntax:'=GCD(A1,B1)', example:'If A1=20 and B1=30, returns 10.' },
{ name:'MROUND', category:'Math', description:'Rounds a number to the nearest multiple.', syntax:'=MROUND(A1,5)', example:'If A1=47, result = 45.' },
{ name:'Volume', category:'Measurement & Geometry', description:'Calculates volume of a box or space.', syntax:'=Length*Width*Height', example:'If dimensions 5×4×3 m, result = 60 m³.' },
{ name:'Area', category:'Measurement & Geometry', description:'Calculates area of a rectangle.', syntax:'=Length*Width', example:'If 20 m × 15 m, result = 300 m².' },
{ name:'Perimeter', category:'Measurement & Geometry', description:'Finds perimeter of rectangle or boundary.', syntax:'=2*(Length+Width)', example:'If 5 m and 3 m sides, result = 16 m.' },
{ name:'PI', category:'Math Constant', description:'Returns the mathematical constant π.', syntax:'=PI()', example:'Useful for circle area formulas (πr²).' },
{ name:'CIRCLEAREA', category:'Geometry', description:'Calculates area of a circle from radius.', syntax:'=PI()*(A1^2)', example:'If radius = 7 m, area ≈ 153.94 m².' },
{ name:'WORKINGDAYS', category:'HR & Operations', description:'Counts number of working days excluding weekends.', syntax:'=NETWORKDAYS(A1,B1)', example:'If A1=1-Mar, B1=31-Mar, returns 23 working days.' },
{ name:'PAYROLL', category:'HR & Payroll', description:'Computes net salary from gross by deductions.', syntax:'=Gross-(PF+TDS+ESI)', example:'If Gross ₹60K, deductions ₹10K, Net = ₹50K.' },
{ name:'ABS', category:'Math', description:'Returns the absolute value of a number.', syntax:'=ABS(A1)', example:'If A1=-25, result = 25.' },
{ name:'ROUND', category:'Math', description:'Rounds a number to specified decimal places.', syntax:'=ROUND(A1,2)', example:'If A1=2.678, result = 2.68.' },
{ name:'SQRT', category:'Math', description:'Returns the square root of a number.', syntax:'=SQRT(A1)', example:'If A1=49, result = 7.' },
{ name:'POWER', category:'Math', description:'Raises a number to a power.', syntax:'=POWER(A1,2)', example:'If A1=6, result = 36.' },

{ name:'SUMSQ', category:'Array Math', description:'Adds up the squares of all given numbers.', syntax:'=SUMSQ(A1:A5)', example:'If A1:A3 = 2,3,4 then =SUMSQ(A1:A3) gives 29.' },
{ name:'DEVSQ', category:'Array Math', description:'Calculates the sum of squared deviations from the mean.', syntax:'=DEVSQ(A1:A10)', example:'Used in risk analysis to measure variation of returns.' },
{ name:'GEOMEAN', category:'Statistical', description:'Finds the geometric mean (useful for growth rates).', syntax:'=GEOMEAN(A1:A5)', example:'If growth rates are 1.10,1.20,1.15, result is 1.15 or 15% average growth.' },
{ name:'HARMEAN', category:'Statistical', description:'Calculates harmonic mean, used in speed and ratios.', syntax:'=HARMEAN(A1:A3)', example:'If speeds are 60, 80, 100 km/h, result = 76.8 km/h average speed.' },
{ name:'FISHER', category:'Statistical', description:'Performs Fisher transformation for correlation data.', syntax:'=FISHER(A1)', example:'Converts correlation coefficient to Z value for testing.' },
{ name:'PHI', category:'Statistical', description:'Calculates the value of the standard normal distribution at a point.', syntax:'=PHI(A1)', example:'Returns probability density for a Z-score.' },
{ name:'LEN', category:'Text Analytics', description:'Counts number of characters in a text.', syntax:'=LEN(A1)', example:'If A1 = “Kolaghat Branch”, result = 15.' },
{ name:'SUBSTITUTE', category:'Text Analytics', description:'Replaces one text with another within a cell.', syntax:'=SUBSTITUTE(A1,"PNB","Punjab National Bank")', example:'Changes “PNB Kolaghat” to “Punjab National Bank Kolaghat”.' },
{ name:'TEXTSPLIT', category:'Text Analytics', description:'Splits text into multiple cells based on a delimiter.', syntax:'=TEXTSPLIT(A1,",")', example:'If A1 = “East,West,North”, result splits into 3 separate cells.' },
{ name:'REGEX.EXTRACT', category:'Text Analytics', description:'Extracts a specific pattern of text using regex.', syntax:'=REGEX.EXTRACT(A1,"[0-9]{6}")', example:'If A1 contains “PIN 721137”, returns 721137.' },
{ name:'STOCKHISTORY', category:'Market & Finance', description:'Fetches historical stock prices directly from web.', syntax:'=STOCKHISTORY("PNB.NS",TODAY()-30,TODAY())', example:'Displays PNB share prices of the last 30 days.' },
{ name:'PRICE', category:'Finance', description:'Calculates bond price per ₹100 face value.', syntax:'=PRICE(Settle,Maturity,Rate,Yld,Redemption,Freq,Basis)', example:'Used to value a bond traded at discount or premium.' },
{ name:'YIELD', category:'Finance', description:'Finds yield on a security that pays periodic interest.', syntax:'=YIELD(Settle,Maturity,Rate,Price,Redemption,Freq,Basis)', example:'If bond pays 8% interest but priced at ₹950, gives yield > 8%.' },
{ name:'DURATION', category:'Finance', description:'Calculates Macaulay duration for a bond.', syntax:'=DURATION(Settle,Maturity,Rate,Yld,Freq,Basis)', example:'Used by bankers to measure interest rate sensitivity of bonds.' },
{ name:'ACCRINT', category:'Finance', description:'Computes accrued interest of a bond up to a given date.', syntax:'=ACCRINT(Issue,Settle,Rate,Par,Freq,Basis)', example:'Shows interest earned by investor before coupon payment.' },
{ name:'PPMT', category:'Loan', description:'Returns principal portion of a loan payment.', syntax:'=PPMT(Rate/12,1,360,-LoanAmt)', example:'Shows principal paid in 1st EMI for a home loan.' },
{ name:'CUMIPMT', category:'Loan', description:'Calculates total interest paid over a number of periods.', syntax:'=CUMIPMT(6%/12,120,500000,-500000,1,12)', example:'Gives total interest paid in first year on ₹5L loan.' },
{ name:'FILTER', category:'Data Analysis', description:'Filters rows based on specific condition.', syntax:'=FILTER(A2:B100,B2:B100>100)', example:'Displays all transactions above ₹100.' },
{ name:'AGGREGATE', category:'Data Analysis', description:'Applies multiple statistical functions while ignoring errors.', syntax:'=AGGREGATE(9,5,A2:A10)', example:'Returns SUM while ignoring hidden or error cells.' },
{ name:'FORECAST.LINEAR', category:'Forecasting', description:'Predicts future values using a linear trend.', syntax:'=FORECAST.LINEAR(NextX,Y_range,X_range)', example:'Estimates next month’s sales based on previous 12 months.' },
{ name:'FORECAST.ETS.SEASONALITY', category:'Forecasting', description:'Detects the seasonality pattern in time series data.', syntax:'=FORECAST.ETS.SEASONALITY(A2:A100,B2:B100,1,1)', example:'Shows 12 for monthly seasonality pattern.' },
{ name:'NETWORKDAYS', category:'Date Intelligence', description:'Counts working days between two dates.', syntax:'=NETWORKDAYS(A1,B1)', example:'If A1=1-Jan and B1=10-Jan, returns 8 (excluding weekends).' },
{ name:'WORKDAY', category:'Date Intelligence', description:'Returns a date after a specified number of workdays.', syntax:'=WORKDAY(A1,5)', example:'If A1=1-Jan, result = 8-Jan (skips weekends).' },
{ name:'YEARFRAC', category:'Date Intelligence', description:'Calculates the fraction of the year between two dates.', syntax:'=YEARFRAC(A1,B1)', example:'If A1=01-Apr and B1=30-Sep, result ≈ 0.5 years.' },
{ name:'EOMONTH', category:'Date Intelligence', description:'Returns the last day of the month after a set number of months.', syntax:'=EOMONTH(A1,2)', example:'If A1=10-Jan, returns 31-Mar.' },
{ name:'FiscalYear', category:'Date Intelligence', description:'Returns fiscal year based on April-March cycle.', syntax:'=IF(MONTH(A1)>=4,YEAR(A1),YEAR(A1)-1)', example:'If A1=15-Mar-2025, result = 2024.' },
{ name:'COVARIANCE.P', category:'Data Science', description:'Calculates population covariance between two datasets.', syntax:'=COVARIANCE.P(A1:A10,B1:B10)', example:'Used to check how two business metrics move together.' },
{ name:'RSQ', category:'Data Science', description:'Calculates R² (coefficient of determination) between two datasets.', syntax:'=RSQ(Y_range,X_range)', example:'If RSQ = 0.85, 85% of sales variation is explained by price.' },
{ name:'VAR.S', category:'Statistical', description:'Calculates sample variance.', syntax:'=VAR.S(A1:A10)', example:'Used to find how much sales figures vary across branches.' },
{ name:'STDEV.S', category:'Statistical', description:'Calculates standard deviation for sample data.', syntax:'=STDEV.S(A1:A10)', example:'Used to measure consistency of deposit growth across branches.' },
{ name:'BINOM.DIST', category:'Probability', description:'Returns the probability of a number of successes in trials.', syntax:'=BINOM.DIST(x,n,p,FALSE)', example:'If 3 successes in 5 trials with p=0.6, gives probability 0.346.' },
{ name:'POISSON.DIST', category:'Probability', description:'Calculates probability of a given number of events.', syntax:'=POISSON.DIST(x,λ,FALSE)', example:'Used to model number of loan applications per day.' },
{ name:'NORM.INV', category:'Probability', description:'Returns the value corresponding to a given probability in a normal distribution.', syntax:'=NORM.INV(p,Mean,SD)', example:'Finds cutoff value for top 5% performers.' },
{ name:'WEIBULL.DIST', category:'Probability', description:'Used for reliability and life data analysis.', syntax:'=WEIBULL.DIST(x,α,β,FALSE)', example:'Used to model failure rates in machine operations.' },
{ name:'EOQ', category:'Operations & Supply Chain', description:'Finds the optimal order quantity minimizing total cost.', syntax:'=SQRT((2*Demand*OrderCost)/HoldCost)', example:'If D=1000, OC=50, HC=2, EOQ = 224 units.' },
{ name:'ReorderPoint', category:'Operations & Supply Chain', description:'Determines when to reorder inventory.', syntax:'=(LeadTime*DemandPerDay)+SafetyStock', example:'If lead time=5 days, demand/day=10, safety=30, result=80 units.' },
{ name:'FillRate', category:'Operations & Supply Chain', description:'Measures how many customer orders are fulfilled.', syntax:'=Delivered/Ordered', example:'If 95 orders delivered out of 100, rate = 95%.' },
{ name:'DATA TABLE', category:'Decision Support', description:'Runs what-if analysis for two variables.', syntax:'=DATA TABLE(RowInput,ColumnInput)', example:'Used to compare EMIs for various rates and tenures.' },
{ name:'GOAL SEEK', category:'Decision Support', description:'Finds the input value to reach a desired output.', syntax:'=GOAL SEEK(TargetCell,InputCell,GoalValue)', example:'Used to find required sales to achieve ₹1L profit.' },
{ name:'SCENARIO MANAGER', category:'Decision Support', description:'Stores and compares multiple sets of assumptions.', syntax:'(Access via Data → What-If Analysis)', example:'Compare optimistic and conservative business projections.' },
{ name:'WEBSERVICE', category:'Web Integration', description:'Imports data from a given API or website.', syntax:'=WEBSERVICE("https://api.exchangerate.host/latest")', example:'Fetches live exchange rates for financial dashboard.' },
{ name:'ENCODEURL', category:'Web Integration', description:'Converts text into a URL-safe format.', syntax:'=ENCODEURL(A1)', example:'If A1 = "PNB India", returns PNB%20India.' },
{ name:'IMAGE', category:'Web Integration', description:'Displays an image from a URL directly in Excel.', syntax:'=IMAGE("https://logo.clearbit.com/pnbindia.in")', example:'Inserts the PNB logo from the web.' },

{ name:'IFERROR', category:'Error Handling', description:'Used to return a custom result instead of an Excel error when a formula fails.', syntax:'=IFERROR(A1/B1,0)', example:'If cell A1=10 and B1=0, =IFERROR(A1/B1,0) returns 0 instead of #DIV/0!.' },
{ name:'IFNA', category:'Error Handling', description:'Displays a custom message when a formula returns #N/A.', syntax:'=IFNA(VLOOKUP(A1,B2:C10,2,0),"Not Found")', example:'If a lookup value is missing, it shows “Not Found” instead of #N/A.' },
{ name:'ERROR.TYPE', category:'Error Handling', description:'Returns a number corresponding to the type of error in a cell.', syntax:'=ERROR.TYPE(A1)', example:'If A1 has #N/A, =ERROR.TYPE(A1) gives 7.' },
{ name:'ISERR', category:'Error Handling', description:'Checks if a cell contains any error except #N/A.', syntax:'=ISERR(A1)', example:'If A1 has #VALUE!, =ISERR(A1) returns TRUE.' },
{ name:'CELL', category:'Information', description:'Returns information about the formatting, location, or contents of a cell.', syntax:'=CELL("address",A1)', example:'Shows the address of A1 as $A$1.' },
{ name:'INFO', category:'Information', description:'Returns information about the current operating environment.', syntax:'=INFO("directory")', example:'Displays the file directory path of the workbook.' },
{ name:'SHEET', category:'Information', description:'Returns the sheet number of the referenced sheet.', syntax:'=SHEET()', example:'If the current sheet is the 3rd sheet, =SHEET() returns 3.' },
{ name:'SHEETS', category:'Information', description:'Counts the total number of sheets in the workbook.', syntax:'=SHEETS()', example:'If there are 5 sheets, it returns 5.' },
{ name:'HYPERLINK', category:'Cross-Sheet & File', description:'Creates a clickable link to another cell, file, or webpage.', syntax:'=HYPERLINK("https://pnbindia.in","PNB Site")', example:'Creates a link titled “PNB Site” to the PNB website.' },
{ name:'INDIRECT', category:'Cross-Sheet & File', description:'Returns the reference specified by a text string.', syntax:'=INDIRECT("Sheet2!A1")', example:'Refers dynamically to cell A1 in Sheet2.' },
{ name:'BreakEvenQty', category:'Business Math', description:'Calculates the units needed to cover total fixed costs.', syntax:'=Fixed_Cost/(Price-Variable_Cost)', example:'If fixed cost ₹50,000, price ₹100, and variable cost ₹60, then output = 1,250 units.' },
{ name:'Contribution', category:'Business Math', description:'Finds the contribution margin per unit.', syntax:'=Price-Variable_Cost', example:'If selling price ₹100 and variable cost ₹60, contribution is ₹40 per unit.' },
{ name:'MarginPercent', category:'Business Math', description:'Calculates profit margin as a percentage of sales.', syntax:'=(Profit/Sales)*100', example:'If profit ₹25,000 and sales ₹1,00,000, margin = 25%.' },
{ name:'MarkUpPercent', category:'Business Math', description:'Shows how much above cost price an item is sold.', syntax:'=(Profit/Cost)*100', example:'If cost ₹200 and profit ₹50, markup = 25%.' },
{ name:'TIMEVALUE', category:'Time & Shift', description:'Converts a time in text format to a serial number.', syntax:'=TIMEVALUE("2:30 PM")', example:'Returns 0.60417 representing 2:30 PM.' },
{ name:'MOD', category:'Time & Shift', description:'Used to calculate duration when time crosses midnight.', syntax:'=MOD(End-Start,1)', example:'If Start=10 PM and End=2 AM, it returns 4 hours.' },
{ name:'ReorderLevel', category:'Inventory', description:'Calculates the minimum stock to reorder.', syntax:'=DailyUse*LeadTime', example:'If daily use = 10 units and lead time = 5 days, reorder level = 50 units.' },
{ name:'SafetyStock', category:'Inventory', description:'Adds buffer stock for uncertainties.', syntax:'=AVERAGE(DailyUse)*2', example:'If average daily use is 20 units, safety stock = 40 units.' },
{ name:'TurnoverRatio', category:'Inventory', description:'Shows how many times inventory is sold and replaced.', syntax:'=COGS/AverageInventory', example:'If COGS = ₹1 Lakh and avg inventory = ₹20 K, ratio = 5.' },
{ name:'WeightedAvgCost', category:'Inventory', description:'Calculates the average cost per unit when prices differ.', syntax:'=SUMPRODUCT(Qty,Price)/SUM(Qty)', example:'If 100 units @₹10 and 200 @₹12, avg = ₹11.33.' },
{ name:'AttendancePercent', category:'HR & Payroll', description:'Calculates employee attendance percentage.', syntax:'=COUNTIF(Range,"P")/COUNTA(Range)', example:'If 20 working days and 18 present, attendance = 90%.' },
{ name:'PFContribution', category:'HR & Payroll', description:'Computes Provident Fund amount from basic pay.', syntax:'=MIN(15000,Basic)*0.12', example:'If basic ₹18 K, PF = ₹1,800.' },
{ name:'Gratuity', category:'HR & Payroll', description:'Estimates gratuity amount payable.', syntax:'=Basic*15/26*Years', example:'If basic ₹20 K and service 10 yrs, gratuity = ₹1.15 Lakh.' },
{ name:'OutstandingLoan', category:'Banking & Loan', description:'Calculates loan balance after certain months.', syntax:'=FV(Rate/12,MonthsPaid,EMI,-LoanAmt)', example:'For ₹5 L loan, 10 years @ 9%, outstanding after 2 yrs ≈ ₹4.5 L.' },
{ name:'RefinanceBenefit', category:'Banking & Loan', description:'Shows EMI difference between two loan offers.', syntax:'=OldEMI-NewEMI', example:'If old EMI ₹25 K, new ₹22 K, benefit ₹3 K/month.' },
{ name:'SharpeRatio', category:'Risk & Portfolio', description:'Measures portfolio performance adjusted for risk.', syntax:'=(Return-RiskFree)/STDEV(Returns)', example:'If portfolio 10%, risk-free 6%, SD 5%, Sharpe = 0.8.' },
{ name:'Beta', category:'Risk & Portfolio', description:'Shows volatility of stock vs market.', syntax:'=COVARIANCE.P(Stock,Market)/VAR.P(Market)', example:'If result = 1.2, stock moves 20% more than market.' },
{ name:'FILTER', category:'Dynamic Filter', description:'Extracts rows that meet specific criteria.', syntax:'=FILTER(A2:B100,B2:B100>100)', example:'Displays all sales over ₹100.' },
{ name:'SORTBY', category:'Dynamic Filter', description:'Sorts filtered data by another column.', syntax:'=SORTBY(A2:B100,B2:B100,-1)', example:'Sorts data by sales in descending order.' },
{ name:'UNIQUE', category:'Dynamic Filter', description:'Removes duplicates from a list.', syntax:'=UNIQUE(A2:A100)', example:'Lists unique branch names from column A.' },
{ name:'TAKE', category:'Dynamic Filter', description:'Returns top N rows from a range.', syntax:'=TAKE(A1:A100,10)', example:'Shows top 10 records from column A.' },
{ name:'INFO("recalc")', category:'Automation', description:'Detects when sheet recalculates.', syntax:'=INFO("recalc")', example:'Helps trigger macro when sheet refreshes.' },
{ name:'FORMULATEXT', category:'Audit', description:'Displays the actual formula used in a cell.', syntax:'=FORMULATEXT(A1)', example:'Shows the text of A1’s formula on screen.' },
{ name:'FORECAST.ETS', category:'Forecast', description:'Predicts future values using exponential smoothing.', syntax:'=FORECAST.ETS(next,A2:A100,B2:B100,1,1)', example:'Forecasts next month’s sales from past data.' },
{ name:'MovingAverage', category:'Forecast', description:'Calculates average over a moving period.', syntax:'=AVERAGE(OFFSET(A1,ROW(A1:A10)-1,0,3,1))', example:'Shows 3-month average sales trend.' },
{ name:'KPI Score', category:'Performance', description:'Evaluates target achievement status.', syntax:'=IF(Ach/Budget>=1,"✔","✘")', example:'Shows ✔ if sales ≥ target, else ✘.' },
{ name:'WeightedScore', category:'Performance', description:'Calculates weighted performance rating.', syntax:'=SUMPRODUCT(Weight,Score)/SUM(Weight)', example:'Computes overall score using weights for each metric.' },
{ name:'RANK.EQ', category:'Performance', description:'Gives the rank of a number in a list.', syntax:'=RANK.EQ(A1,A:A)', example:'Ranks a branch by its deposit within circle.' },
{ name:'HighlightTop5', category:'Conditional Formatting', description:'Formula to highlight top 5 values.', syntax:'=RANK(A1,$A$1:$A$100)<=5', example:'Used with conditional formatting to mark top 5 performers.' },
{ name:'COUNTIF', category:'Conditional Formatting', description:'Counts cells that meet a condition.', syntax:'=COUNTIF(A1:A100,">100")', example:'Counts entries above ₹100 crore target.' },
{ name:'RAND', category:'Simulation', description:'Generates a random number between 0 and 1.', syntax:'=RAND()', example:'Each recalculation gives a new decimal value like 0.48.' },
{ name:'RANDBETWEEN', category:'Simulation', description:'Generates a random integer in a range.', syntax:'=RANDBETWEEN(1,100)', example:'Returns a random whole number between 1 and 100.' },
{ name:'ExpectedValue', category:'Simulation', description:'Calculates expected outcome from probabilities.', syntax:'=SUMPRODUCT(Probabilities,Outcomes)', example:'If 0.6×100 + 0.4×50, result = 80.' },
{ name:'CORREL', category:'Data Science', description:'Finds correlation between two datasets.', syntax:'=CORREL(A:A,B:B)', example:'Checks how sales move with advertising spend.' },
{ name:'LINEST', category:'Data Science', description:'Returns regression line statistics.', syntax:'=LINEST(Y_range,X_range,TRUE,TRUE)', example:'Estimates relationship between sales and price.' },
{ name:'WEBSERVICE', category:'Web Integration', description:'Fetches data from a web URL.', syntax:'=WEBSERVICE("https://api.exchangerate.host/latest")', example:'Pulls live currency exchange rates into Excel.' },
{ name:'FILTERXML', category:'Web Integration', description:'Extracts XML data from a web service.', syntax:'=FILTERXML(A1,"//USD")', example:'Extracts USD rate from the XML response.' },
        ];
// Auto-assign a difficulty level to each formula based on category

const BASIC_CATS = new Set([
  'Core Excel Starter Kit', 'Basics & Logic', 'Text & Formatting', 'Count & Frequency',
  'Math & Rounding', 'Date & Time', 'Information', 'Math & Trig', 'Text', 'Financial'
]);

const INTERMEDIATE_CATS = new Set([
  'Lookup & Reference', 'Advanced Lookup', 'Conditional Math', 'Database & Information',
  'Dashboard & Sheet Info', 'Array & Filter', 'Data Validation & List Automation',
  'Dynamic Array & Lambda', 'Data Analysis', 'Forecasting', 'Power Query / Data Model'
]);

const ADVANCED_CATS = new Set([
  'Probability & Statistical Advanced', 'Advanced Statistical & Analytical Tools',
  'Array Math & Matrix Operations', 'Excel + AI Integration', 'Simulation',
  'AI & Copilot', 'Matrix & Geometry', 'Risk & Portfolio', 'Scenario Analysis',
  'Web Integration', 'Dynamic Data Cleaning & Transformation'
]);

function categoryToLevel(cat) {
  if (BASIC_CATS.has(cat)) return 'Basic';
  if (INTERMEDIATE_CATS.has(cat)) return 'Intermediate';
  if (ADVANCED_CATS.has(cat)) return 'Advanced';
  return 'Intermediate'; // fallback
}

// 🔄 Assign level to each formula if missing
formulas.forEach(f => {
  if (!f.level) f.level = categoryToLevel(f.category);
});


        const shortcuts = [
            // General
            { name: 'New Workbook', keys: 'Ctrl + N', category: 'General', description: 'A: Press these keys. B: A new, blank Excel file opens instantly.' },
            { name: 'Save Workbook', keys: 'Ctrl + S', category: 'General', description: 'A: Press these keys. B: It saves your current work so you don\'t lose it.' },
            { name: 'Format Cells', keys: 'Ctrl + 1', category: 'General', description: 'A: Select a cell. B: Press these keys. C: A menu opens to change font, color, number type, and more.' },
            { name: 'Select All', keys: 'Ctrl + A', category: 'General', description: 'A: Click any cell in your data. B: Press these keys. C: It selects your entire block of data.' },
            { name: 'Select Row', keys: 'Shift + Space', category: 'General', description: 'A: Click any cell. B: Press these keys. C: The entire row is selected.' },
            { name: 'Select Column', keys: 'Ctrl + Space', category: 'General', description: 'A: Click any cell. B: Press these keys. C: The entire column is selected.' },
            { name: 'Move to Edge', keys: 'Ctrl + Arrow Key', category: 'General', description: 'A: Hold Ctrl. B: Press an arrow key. C: You jump to the very end of your data in that direction.' },
            { name: 'Go to A1', keys: 'Ctrl + Home', category: 'General', 'description': 'A: Press these keys. B: You instantly jump to the very first cell, A1.' },
            { name: 'Toggle Formulas/Values', keys: 'Ctrl + `', category: 'General', description: 'A: Press these keys. B: See all the formulas in your sheet instead of the results. Press again to switch back.' },
            { name: 'AutoSum', keys: 'Alt + =', category: 'General', description: 'A: Click below a column of numbers. B: Press these keys. C: Excel automatically writes a SUM formula for you.' },
            { name: 'Toggle Ref Type (A1/$A$1)', keys: 'F4', category: 'General', description: 'A: When writing a formula, click a cell reference like A1. B: Press F4. C: It adds dollar signs ($A$1) to lock the reference.' },
            { name: 'Insert Date', keys: 'Ctrl + ;', category: 'General', description: 'A: Click a cell. B: Press these keys. C: Today\'s date instantly appears.' },
            
            // Power
            { name: 'Paste Special', keys: 'Ctrl + Alt + V', category: 'Power', description: 'A: Copy a cell. B: Press these keys. C: A menu appears to paste just the value, format, or formula, which is incredibly useful.' },
            { name: 'Flash Fill', keys: 'Ctrl + E', category: 'Power', description: 'A: Type an example of what you want in the column next to your data (e.g., first names from a full name column). B: Press these keys. C: Excel figures out the pattern and fills the rest for you like magic.' },
            { name: 'Create Table', keys: 'Ctrl + T', category: 'Power', description: 'A: Click your data. B: Press these keys. C: Your data is turned into a smart table with sorting, filtering, and auto-expanding formulas.' },
            { name: 'Quick Analysis', keys: 'Ctrl + Q', category: 'Power', description: 'A: Select your data. B: Press these keys. C: A small icon appears letting you instantly add charts, totals, conditional formatting, and sparklines.' },
            { name: 'Go To Special', keys: 'F5 or Ctrl + G, then Alt + S', category: 'Power', description: 'A: Press F5, then Alt+S. B: A menu appears to select all cells that are blank, have formulas, or contain errors. C: An amazing tool for cleaning up messy data.' },
        ];

// Give every item a default level so filtering works
(formulas || []).forEach(it => { if (!it.level) it.level = 'Basic'; });
(shortcuts || []).forEach(it => { if (!it.level) it.level = 'Basic'; });

        const synonymMap = {
            // This map enhances search by linking related terms.
            'add': ['add', 'addition', 'sum', 'total', 'plus', 'calculate'],
            'addition': ['add', 'addition', 'sum', 'total', 'plus', 'calculate'],
            'sum': ['add', 'addition', 'sum', 'total', 'plus', 'calculate'],
            'subtract': ['subtract', 'minus', 'difference', 'less', 'remove'],
            'minus': ['subtract', 'minus', 'difference', 'less', 'remove'],
            'multiply': ['multiply', 'product', 'times'],
            'product': ['multiply', 'product', 'times'],
            'divide': ['divide', 'division', 'quotient'],
            'average': ['average', 'mean', 'avg'],
            'mean': ['average', 'mean', 'avg'],
            'count': ['count', 'tally', 'how many'],
            'find': ['find', 'lookup', 'search', 'get', 'match', 'locate', 'retrieve'],
            'lookup': ['find', 'lookup', 'search', 'get', 'match', 'locate', 'retrieve'],
            'get': ['find', 'lookup', 'search', 'get', 'match', 'locate', 'retrieve'],
            'join': ['join', 'concat', 'combine', 'merge', 'textjoin', 'append'],
            'combine': ['join', 'concat', 'combine', 'merge', 'textjoin', 'append'],
            'text': ['text', 'string', 'character', 'word'],
            'string': ['text', 'string', 'character', 'word'],
            'flip': ['flip', 'transpose', 'reverse', 'rotate'],
            'transpose': ['flip', 'transpose', 'reverse', 'rotate'],
            'clean': ['clean', 'trim', 'remove spaces', 'tidy'],
            'if': ['if', 'condition', 'logical', 'test', 'check'],
            'condition': ['if', 'condition', 'logical', 'test', 'check'],
            'date': ['date', 'time', 'day', 'month', 'year', 'now', 'today'],
            'time': ['date', 'time', 'hour', 'minute', 'second', 'now'],
            'round': ['round', 'decimal', 'integer'],
            'error': ['error', 'mistake', 'problem', 'issue'],
            'random': ['random', 'rand', 'dice roll'],
            'list': ['list', 'array', 'range', 'data'],
            'array': ['list', 'array', 'range', 'data'],
            'table': ['table', 'database', 'pivot'],
            'big': ['big', 'large', 'max', 'maximum', 'highest'],
            'large': ['big', 'large', 'max', 'maximum', 'highest'],
            'small': ['small', 'min', 'minimum', 'lowest', 'little'],
            'min': ['small', 'min', 'minimum', 'lowest', 'little'],
        };

        document.addEventListener('DOMContentLoaded', () => {
            // --- EmailJS Initialization ---
            // IMPORTANT: To enable email features, create an account at https://www.emailjs.com/
            // and enter your credentials in the constants below. This is a secure way to manage keys.
            const EMAILJS_USER_ID = ''; // Your EmailJS User ID (Public Key)
            const EMAILJS_SERVICE_ID = ''; // Your EmailJS Service ID
            const EMAILJS_WELCOME_TEMPLATE_ID = ''; // Your "Welcome Email" Template ID
            const EMAILJS_CONTACT_TEMPLATE_ID = ''; // Your "Contact Form" Template ID

            (function(){
                // The init function should only be called if the user ID is provided.
                if (EMAILJS_USER_ID) {
                    emailjs.init(EMAILJS_USER_ID); 
                }
            })();

            const contentGrid = document.getElementById('content-grid');
            const searchInput = document.getElementById('searchInput');
            const tabs = document.querySelectorAll('.tab-btn');
            const categoryFiltersContainer = document.getElementById('category-filters');
            const noResults = document.getElementById('no-results');
            const copiedToast = document.getElementById('copied-toast');
            const alphabetFiltersContainer = document.getElementById('alphabet-filters');
            const activeTabSlider = document.getElementById('active-tab-slider');
            
            // Modal elements
            const exampleModal = document.getElementById('exampleModal');
            const closeModalBtn = document.getElementById('closeModal');
            const modalTitle = document.getElementById('modalTitle');
            const modalSyntax = document.getElementById('modalSyntax');
            const modalExample = document.getElementById('modalExample');
            const greetingElement = document.getElementById('greeting');
            const timeDisplay = document.getElementById('time-display');
            const dateDisplay = document.getElementById('date-display');

            // --- Auth Elements ---
            const loginModal = document.getElementById('loginModal');
            const closeLoginModal = document.getElementById('closeLoginModal');
            const loginMobileInput = document.getElementById('loginMobileInput');
            const loginPinInput = document.getElementById('loginPinInput');
            const submitLoginBtn = document.getElementById('submitLoginBtn');
            const loginAuthError = document.getElementById('loginAuthError');

            const signupModal = document.getElementById('signupModal');
            const closeSignupModal = document.getElementById('closeSignupModal');
            const signupNameInput = document.getElementById('signupNameInput');
            const signupEmailInput = document.getElementById('signupEmailInput');
            const signupMobileInput = document.getElementById('signupMobileInput');
            const signupPincodeInput = document.getElementById('signupPincodeInput');
            const signupPinInput = document.getElementById('signupPinInput');
            const submitSignupBtn = document.getElementById('submitSignupBtn');
            const signupAuthError = document.getElementById('signupAuthError');
            const termsAgreeCheckbox = document.getElementById('terms-agree-checkbox');
            const privacyPolicyLinkSignup = document.getElementById('privacy-policy-link-signup');
            const termsOfUseLinkSignup = document.getElementById('terms-of-use-link-signup');

            const loginBtn = document.getElementById('loginBtn');
            const signupBtn = document.getElementById('signupBtn');
            const logoutBtn = document.getElementById('logoutBtn');
            const myProfileBtn = document.getElementById('myProfileBtn');
            const menuToggleBtn = document.getElementById('menu-toggle-btn');
            const menuDropdown = document.getElementById('menu-dropdown');
            
            const profileModal = document.getElementById('profileModal');
            const closeProfileModal = document.getElementById('closeProfileModal');
            const profileDetails = document.getElementById('profileDetails');

            const welcomeModal = document.getElementById('welcomeModal');
            const welcomeMessage = document.getElementById('welcomeMessage');
            const closeWelcomeModal = document.getElementById('closeWelcomeModal');
            
            // --- Legal Modal Elements ---
            const privacyModal = document.getElementById('privacyModal');
            const closePrivacyModal = document.getElementById('closePrivacyModal');
            const privacyPolicyLink = document.getElementById('privacy-policy-link');
            const termsModal = document.getElementById('termsModal');
            const closeTermsModal = document.getElementById('closeTermsModal');
            const termsOfUseLink = document.getElementById('terms-of-use-link');


            // --- Footer Elements ---
            document.getElementById('copyright-year').textContent = new Date().getFullYear();
            const contactForm = document.getElementById('contact-form');
            const contactResponse = document.getElementById('contact-response');
            const submitContactBtn = document.getElementById('submit-contact-btn');
            const contactConsentCheckbox = document.getElementById('contact-consent-checkbox');


            // --- Greeting & Auth Logic ---
            let timeOfDayGreeting = '';
            const currentHour = new Date().getHours();
            if (currentHour < 12) timeOfDayGreeting = 'Good Morning';
            else if (currentHour < 17) timeOfDayGreeting = 'Good Afternoon';
            else timeOfDayGreeting = 'Good Evening';
            
            // Simple hashing function for PIN security
            const hashPIN = (pin) => {
                const salt = "ExcelVault_Salt_!@#$"; // A static salt for basic security
                let hash = 0;
                const saltedPin = salt + pin + salt;
                for (let i = 0; i < saltedPin.length; i++) {
                    const char = saltedPin.charCodeAt(i);
                    hash = ((hash << 5) - hash) + char;
                    hash |= 0; // Convert to 32bit integer
                }
                return hash.toString();
            };

            const updateUserUI = (user = null) => {
                const firstName = user ? user.name.split(' ')[0] : "Amit";
                const capitalizedFirstName = firstName.charAt(0).toUpperCase() + firstName.slice(1).toLowerCase();

                greetingElement.textContent = `${timeOfDayGreeting}, ${capitalizedFirstName}!`;

                if (user) {
                    myProfileBtn.classList.remove('hidden');
                    logoutBtn.classList.remove('hidden');
                    loginBtn.classList.add('hidden');
                    signupBtn.classList.add('hidden');
                } else {
                    myProfileBtn.classList.add('hidden');
                    logoutBtn.classList.add('hidden');
                    loginBtn.classList.remove('hidden');
                    signupBtn.classList.remove('hidden');
                }
            };
            
            const showLoginModal = () => {
                loginAuthError.classList.add('hidden');
                loginModal.classList.remove('hidden');
                loginModal.classList.add('flex');
                setTimeout(() => loginMobileInput.focus(), 50);
            };

            const hideLoginModal = () => {
                loginModal.classList.add('hidden');
                loginModal.classList.remove('flex');
            };

            const showSignupModal = () => {
                signupAuthError.classList.add('hidden');
                signupModal.classList.remove('hidden');
                signupModal.classList.add('flex');
                setTimeout(() => signupNameInput.focus(), 50);
            };

            const hideSignupModal = () => {
                signupModal.classList.add('hidden');
                signupModal.classList.remove('flex');
            };

            const showWelcomeModal = (userName) => {
                const firstName = userName.split(' ')[0];
                welcomeMessage.textContent = `Welcome, ${firstName}!`;
                welcomeModal.classList.remove('hidden');
                welcomeModal.classList.add('flex');
            };

            const hideWelcomeModal = () => {
                welcomeModal.classList.add('hidden');
                welcomeModal.classList.remove('flex');
            };
            
            const sendWelcomeEmail = (name, email) => {
                if (!EMAILJS_SERVICE_ID || !EMAILJS_WELCOME_TEMPLATE_ID) {
                    console.warn("EmailJS Service ID or Welcome Template ID is not configured. Welcome email not sent.");
                    return;
                }

                const templateParams = {
                    to_name: name.split(' ')[0], // Send first name
                    to_email: email,
                    from_name: 'Excel Speed & Formula Vault'
                };

                emailjs.send(EMAILJS_SERVICE_ID, EMAILJS_WELCOME_TEMPLATE_ID, templateParams)
                    .then(function(response) {
                       console.log('SUCCESS! Welcome email sent.', response.status, response.text);
                    }, function(error) {
                       console.log('FAILED... Could not send welcome email.', error);
                    });
            };

            const handleSignup = () => {
                signupAuthError.classList.add('hidden');
                const name = signupNameInput.value.trim();
                const email = signupEmailInput.value.trim();
                const mobile = signupMobileInput.value.trim();
                const pincode = signupPincodeInput.value.trim();
                const pin = signupPinInput.value.trim();

                if (!name || !mobile || !pin || !email) {
                    signupAuthError.textContent = 'Name, Email, Mobile, and PIN are required.';
                    signupAuthError.classList.remove('hidden');
                    return;
                }
                 if (!/^\S+@\S+\.\S+$/.test(email)) {
                    signupAuthError.textContent = 'Please enter a valid email address.';
                    signupAuthError.classList.remove('hidden');
                    return;
                }
                if (!/^\d{4}$/.test(pin)) {
                    signupAuthError.textContent = 'Entry PIN must be exactly 4 digits.';
                    signupAuthError.classList.remove('hidden');
                    return;
                }

                let users = JSON.parse(localStorage.getItem('excelVaultUsers')) || [];
                if (users.find(u => u.mobile === mobile)) {
                    signupAuthError.textContent = 'This mobile number is already registered.';
                    signupAuthError.classList.remove('hidden');
                    return;
                }
                
                const hashedPin = hashPIN(pin); // Hash the PIN
                const newUser = { name, email, mobile, pincode, pin: hashedPin }; // Store the hashed PIN
                users.push(newUser);
                localStorage.setItem('excelVaultUsers', JSON.stringify(users));
                localStorage.setItem('excelVaultCurrentUser', JSON.stringify(newUser));

                updateUserUI(newUser);
                hideSignupModal();
                showWelcomeModal(newUser.name);
                sendWelcomeEmail(newUser.name, newUser.email); // Send email after successful signup
            };

            const handleLogin = () => {
                loginAuthError.classList.add('hidden');
                const mobile = loginMobileInput.value;
                const pin = loginPinInput.value;
                
                const hashedPin = hashPIN(pin); // Hash the entered PIN for comparison
                let users = JSON.parse(localStorage.getItem('excelVaultUsers')) || [];
                const foundUser = users.find(u => u.mobile === mobile && u.pin === hashedPin); // Compare hashes

                if (foundUser) {
                    localStorage.setItem('excelVaultCurrentUser', JSON.stringify(foundUser));
                    updateUserUI(foundUser);
                    hideLoginModal();
                    loginMobileInput.value = '';
                    loginPinInput.value = '';
                } else {
                    loginAuthError.textContent = 'Invalid mobile number or PIN.';
                    loginAuthError.classList.remove('hidden');
                }
            };

            const handleLogout = () => {
                localStorage.removeItem('excelVaultCurrentUser');
                updateUserUI(null);
		document.getElementById('greeting').textContent = `${timeOfDayGreeting}, Guest!`;
            };

            loginBtn.addEventListener('click', () => { menuDropdown.classList.add('hidden'); showLoginModal(); });
            signupBtn.addEventListener('click', () => { menuDropdown.classList.add('hidden'); showSignupModal(); });
            logoutBtn.addEventListener('click', () => { menuDropdown.classList.add('hidden'); handleLogout(); });

            submitLoginBtn.addEventListener('click', handleLogin);
            submitSignupBtn.addEventListener('click', handleSignup);

            closeLoginModal.addEventListener('click', hideLoginModal);
            loginModal.addEventListener('click', (e) => { if (e.target === loginModal) hideLoginModal(); });
            loginPinInput.addEventListener('keyup', (e) => { if (e.key === 'Enter') handleLogin(); });

            closeSignupModal.addEventListener('click', hideSignupModal);
            signupModal.addEventListener('click', (e) => { if (e.target === signupModal) hideSignupModal(); });
            signupPinInput.addEventListener('keyup', (e) => { if (e.key === 'Enter') handleSignup(); });
            
            // Logic for signup consent checkbox
            termsAgreeCheckbox.addEventListener('change', () => {
                submitSignupBtn.disabled = !termsAgreeCheckbox.checked;
            });
            privacyPolicyLinkSignup.addEventListener('click', (e) => { e.preventDefault(); showPrivacyModal(); });
            termsOfUseLinkSignup.addEventListener('click', (e) => { e.preventDefault(); showTermsModal(); });


            // Welcome Modal Logic
            closeWelcomeModal.addEventListener('click', hideWelcomeModal);
            welcomeModal.addEventListener('click', (e) => { if (e.target === welcomeModal) hideWelcomeModal(); });

            // Menu Dropdown Logic
            menuToggleBtn.addEventListener('click', (e) => {
                e.stopPropagation();
                const isExpanded = menuToggleBtn.getAttribute('aria-expanded') === 'true';
                menuToggleBtn.setAttribute('aria-expanded', !isExpanded);
                menuDropdown.classList.toggle('hidden');
            });
            window.addEventListener('click', () => {
                if (!menuDropdown.classList.contains('hidden')) {
                    menuDropdown.classList.add('hidden');
                    menuToggleBtn.setAttribute('aria-expanded', 'false');
                }
            });
            menuDropdown.addEventListener('click', e => e.stopPropagation());

            // Profile Modal Logic
            const hideProfileModal = () => {
                profileModal.classList.add('hidden');
                profileModal.classList.remove('flex');
            };

            myProfileBtn.addEventListener('click', (e) => {
                e.preventDefault();
                menuDropdown.classList.add('hidden');
                menuToggleBtn.setAttribute('aria-expanded', 'false');
                const currentUser = JSON.parse(localStorage.getItem('excelVaultCurrentUser'));
                if (currentUser) {
                    profileDetails.innerHTML = ''; 

                    const createDetailLine = (label, value) => {
                        const p = document.createElement('p');
                        const strong = document.createElement('strong');
                        strong.className = 'font-semibold';
                        strong.textContent = `${label}: `;
                        p.appendChild(strong);
                        p.appendChild(document.createTextNode(value || 'Not provided'));
                        profileDetails.appendChild(p);
                    };

                    createDetailLine('Name', currentUser.name);
                    createDetailLine('Email', currentUser.email);
                    createDetailLine('Mobile', currentUser.mobile);
                    createDetailLine('Postal Code', currentUser.pincode);
                    
                    profileModal.classList.remove('hidden');
                    profileModal.classList.add('flex');
                }
            });
            closeProfileModal.addEventListener('click', hideProfileModal);
            profileModal.addEventListener('click', (e) => { if (e.target === profileModal) hideProfileModal(); });


            // --- End Greeting & Auth Logic ---
            
            // --- Legal Modals Logic ---
            const showPrivacyModal = () => { privacyModal.classList.remove('hidden'); privacyModal.classList.add('flex'); };
            const hidePrivacyModal = () => { privacyModal.classList.add('hidden'); privacyModal.classList.remove('flex'); };
            const showTermsModal = () => { termsModal.classList.remove('hidden'); termsModal.classList.add('flex'); };
            const hideTermsModal = () => { termsModal.classList.add('hidden'); termsModal.classList.remove('flex'); };
            
            privacyPolicyLink.addEventListener('click', (e) => { e.preventDefault(); showPrivacyModal(); });
            closePrivacyModal.addEventListener('click', hidePrivacyModal);
            privacyModal.addEventListener('click', (e) => { if (e.target === privacyModal) hidePrivacyModal(); });
            
            termsOfUseLink.addEventListener('click', (e) => { e.preventDefault(); showTermsModal(); });
            closeTermsModal.addEventListener('click', hideTermsModal);
            termsModal.addEventListener('click', (e) => { if (e.target === termsModal) hideTermsModal(); });


            // --- Contact Form Logic ---
             contactConsentCheckbox.addEventListener('change', () => {
                submitContactBtn.disabled = !contactConsentCheckbox.checked;
            });

            contactForm.addEventListener('submit', function(event) {
                event.preventDefault();
                submitContactBtn.textContent = 'Sending...';
                submitContactBtn.disabled = true;
                
                if (!EMAILJS_SERVICE_ID || !EMAILJS_CONTACT_TEMPLATE_ID) {
                    console.warn("EmailJS is not configured for the contact form. Simulating success.");
                    contactResponse.textContent = "Thank you! Your message has been received.";
                    contactResponse.classList.add('text-green-500');
                    contactResponse.classList.remove('text-red-500');
                    contactForm.reset();
                    contactConsentCheckbox.checked = false;
                    submitContactBtn.textContent = 'Send Message';
                    submitContactBtn.disabled = true;
                    return;
                }

                emailjs.sendForm(EMAILJS_SERVICE_ID, EMAILJS_CONTACT_TEMPLATE_ID, this)
                    .then(() => {
                        contactResponse.textContent = 'Thank you! Your message has been sent successfully.';
                        contactResponse.classList.add('text-green-500');
                        contactResponse.classList.remove('text-red-500');
                        contactForm.reset();
                    }, (err) => {
                        contactResponse.textContent = 'Sorry, something went wrong. Please try again.';
                        contactResponse.classList.add('text-red-500');
                        contactResponse.classList.remove('text-green-500');
                        console.error('Failed to send email. Error: ', JSON.stringify(err));
                    })
                    .finally(() => {
                        submitContactBtn.textContent = 'Send Message';
                        submitContactBtn.disabled = true;
                        contactConsentCheckbox.checked = false;
                    });
            });

            // --- Date & Time Logic ---
            const updateDateTime = () => {
                const now = new Date();
                const timeOptions = { hour: '2-digit', minute: '2-digit', hour12: true };
                const dateOptions = { weekday: 'long', year: 'numeric', month: 'short', day: 'numeric' };

                timeDisplay.textContent = now.toLocaleTimeString('en-US', timeOptions);
                dateDisplay.textContent = now.toLocaleDateString('en-US', dateOptions);
            };
            
            updateDateTime();
            setInterval(updateDateTime, 1000); // Update every second
            // --- End Date & Time Logic ---

            let currentTab = 'formulas';
            let activeCategory = 'All';
            let activeLetter = 'All';
	    // Level filter state (default = All)
let activeLevel = 'All';

// Grab buttons
const btnLevelAll          = document.getElementById('btnLevelAll');
const btnLevelBasic        = document.getElementById('btnLevelBasic');
const btnLevelIntermediate = document.getElementById('btnLevelIntermediate');
const btnLevelAdvanced     = document.getElementById('btnLevelAdvanced');

['btnLevelAll', 'btnLevelBasic', 'btnLevelIntermediate', 'btnLevelAdvanced'].forEach(id => {
  const btn = document.getElementById(id);
  btn?.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' || e.key === ' ') {
      e.preventDefault();

      btn.click(); // trigger the same as mouse click
    }
  });
});


// Function to mark active button + re-render
function setActiveLevel(level) {
  activeLevel = level;
  const buttons = [btnLevelAll, btnLevelBasic, btnLevelIntermediate, btnLevelAdvanced];
  buttons.forEach(b => {
    if (!b) return;
    b.classList.remove('bg-blue-600','text-white','shadow');
    b.classList.add('bg-transparent');
  });
  const map = {All:btnLevelAll,Basic:btnLevelBasic,Intermediate:btnLevelIntermediate,Advanced:btnLevelAdvanced};
  if (map[level]) {
    map[level].classList.remove('bg-transparent');
    map[level].classList.add('bg-blue-600','text-white','shadow');
  }
  if (typeof renderContent === 'function') renderContent();
}

// Wire click events
btnLevelAll.addEventListener('click',          () => setActiveLevel('All'));
btnLevelBasic.addEventListener('click',        () => setActiveLevel('Basic'));
btnLevelIntermediate.addEventListener('click', () => setActiveLevel('Intermediate'));
btnLevelAdvanced.addEventListener('click',     () => setActiveLevel('Advanced'));
    
const btnReset = document.getElementById('btnResetFilters');

btnReset?.addEventListener('click', () => {
  // Reset all filter states
  currentTab = 'formulas';
  activeCategory = 'All';
  activeLevel = 'All';
  activeLetter = 'All';
  searchInput.value = '';

// Highlight the Formulas tab
document.querySelectorAll('[data-tab]').forEach(btn => {
  btn.classList.remove('tab-active'); // remove highlight from all
});
document.querySelector('[data-tab="formulas"]')?.classList.add('tab-active'); // add to Formulas tab


  // Reset tab styles
  tabs.forEach(t => {
    t.classList.remove('tab-active');
    if (t.dataset.tab === 'formulas') {
      t.classList.add('tab-active');
    }
  });

  // Reset level button highlights
  if (typeof setActiveLevel === 'function') {
    setActiveLevel('All');
  }

  // Reset category UI if needed
  if (typeof updateCategoryUI === 'function') {
    updateCategoryUI('All');
  }

  // Reset alphabet highlights (if used)
  document.querySelectorAll('.alphabet-filter button').forEach(btn => {
    btn.classList.remove('bg-blue-600', 'text-white');
    if (btn.dataset.letter === 'All') {
      btn.classList.add('bg-blue-600', 'text-white');
    }
  });

  // Scroll to top (optional)
  window.scrollTo({ top: 0, behavior: 'smooth' });

  // Re-render content
  renderContent();
});


// Allow external UI (dropdown) to change the active category safely
window.setCategory = function(cat) {
  // cat must be 'All' or one of your category labels
  activeCategory = cat;
  // Update the button label if it exists
  const label = document.getElementById('selectedCategoryText');
  if (label) label.textContent = (cat === 'All' ? 'All Categories' : cat);
  // Re-render with new category
  if (typeof renderContent === 'function') renderContent();
  // (Optional) debug:
  // console.log('[setCategory] ->', cat);
};
            
            // Load favorites from localStorage
            let favoriteFormulas = JSON.parse(localStorage.getItem('favoriteFormulas')) || [];
            let favoriteShortcuts = JSON.parse(localStorage.getItem('favoriteShortcuts')) || [];

            formulas.forEach(f => f.favorite = favoriteFormulas.includes(f.name));
            shortcuts.forEach(s => s.favorite = favoriteShortcuts.includes(s.name));

	    // --- WhatsApp share helper ---
function buildWhatsAppShareURL(item) {
  const valueToShare = item.syntax || item.keys || '';
  const text = `${item.name} ${valueToShare}\nTry it: ${location.href}`;
  return `https://wa.me/?text=${encodeURIComponent(text)}`;
}


            const createCard = (item, type) => {
                const card = document.createElement('div');
                card.className = 'formula-card bg-white dark:bg-gray-800 rounded-lg shadow-lg p-5 flex flex-col';
                
                const copyValue = item.syntax || item.keys;
                const isFavorited = item.favorite ? 'favorited' : '';
                
                card.innerHTML = `
                    <div class="flex-grow">
                        <div class="flex justify-between items-start gap-2">
                             <h3 class="text-xl font-bold text-gray-900 dark:text-white">${item.name}</h3>
                             <button class="favorite-btn p-1 -mt-1 -mr-1 text-gray-400 hover:text-yellow-400 transition-colors ${isFavorited}" data-name="${item.name}" data-type="${type}">
                                 <svg xmlns="http://www.w3.org/2000/svg" class="w-6 h-6 transition-transform" fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2"><path stroke-linecap="round" stroke-linejoin="round" d="M11.049 2.927c.3-.921 1.603-.921 1.902 0l1.519 4.674a1 1 0 00.95.69h4.915c.969 0 1.371 1.24.588 1.81l-3.976 2.888a1 1 0 00-.363 1.118l1.518 4.674c.3.922-.755 1.688-1.538 1.118l-3.976-2.888a1 1 0 00-1.176 0l-3.976 2.888c-.783.57-1.838-.197-1.538-1.118l1.518-4.674a1 1 0 00-.363-1.118l-3.976-2.888c-.783-.57-.38-1.81.588-1.81h4.914a1 1 0 00.951-.69l1.519-4.674z" /></svg>
                             </button>
                        </div>
                        ${item.category ? `<span class="text-xs font-semibold bg-blue-100 text-blue-800 dark:bg-blue-900 dark:text-blue-200 px-2 py-1 rounded-full inline-block mt-1">${item.category}</span>` : ''}
                        <p class="text-gray-600 dark:text-gray-400 mt-2">${item.description}</p>
                    </div>
                    <div class="mt-4 pt-4 border-t border-gray-200 dark:border-gray-700">
                        <p class="font-mono text-sm bg-gray-100 dark:bg-gray-700 p-2 rounded break-words">${copyValue}</p>
                        
                        <!-- Button Group -->
<div class="flex items-center gap-2 mt-3">
  ${item.example ? `<button class="example-btn w-full px-4 py-2 bg-gray-200 dark:bg-gray-600 text-gray-800 dark:text-gray-200 rounded-lg hover:bg-gray-300 dark:hover:bg-gray-500 focus:outline-none focus:ring-2 focus:ring-gray-400 transition-colors">Example</button>` : '<div class="w-full"></div>'}
  <button class="copy-btn w-full px-4 py-2 bg-blue-500 text-white rounded-lg hover:bg-blue-600 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:ring-opacity-50 transition-colors" data-copy="${copyValue}">Copy</button>
</div>
</div>
                `;

                card.querySelector('.copy-btn').addEventListener('click', e => copyToClipboard(e.target.getAttribute('data-copy')));
                card.querySelector('.favorite-btn').addEventListener('click', toggleFavorite);
               
 // --- WhatsApp Icon Button ---
const btnGroup = card.querySelector('.flex.items-center.gap-2.mt-3');
if (btnGroup) {
  const shareLink = document.createElement('a');
  shareLink.href = buildWhatsAppShareURL(item);
  shareLink.target = '_blank';
  shareLink.rel = 'noopener';
  shareLink.className = 'w-10 h-10 bg-green-500 rounded-lg flex items-center justify-center hover:bg-green-600 focus:outline-none focus:ring-2 focus:ring-green-500 focus:ring-opacity-50 transition-colors';
  shareLink.setAttribute('aria-label', 'Share on WhatsApp');
  shareLink.innerHTML = `
    <svg xmlns="http://www.w3.org/2000/svg" class="w-5 h-5 text-white" viewBox="0 0 24 24" fill="currentColor" aria-hidden="true">
      <path d="M20.52 3.48A11.93 11.93 0 0 0 12.06 0C5.5 0 .2 5.29.2 11.82c0 2.08.55 4.11 1.6 5.9L0 24l6.43-1.74a11.77 11.77 0 0 0 5.63 1.43h.01c6.55 0 11.86-5.3 11.86-11.84 0-3.17-1.23-6.16-3.41-8.37ZM12.07 21.3h-.01a9.5 9.5 0 0 1-4.85-1.32l-.35-.21-3.82 1.03 1.02-3.72-.22-.38a9.47 9.47 0 0 1-1.43-5.09c0-5.24 4.27-9.5 9.52-9.5 2.54 0 4.92.99 6.71 2.78a9.43 9.43 0 0 1 2.8 6.7c0 5.24-4.28 9.51-9.57 9.51Zm5.47-7.1c-.3-.16-1.77-.88-2.04-.98-.27-.1-.47-.16-.68.16-.2.31-.78.97-.96 1.17-.18.2-.35.22-.65.08-.3-.15-1.28-.47-2.44-1.49-.9-.78-1.51-1.75-1.69-2.05-.18-.3-.02-.46.13-.62.13-.13.3-.35.44-.52.15-.17.2-.31.3-.52.1-.2.05-.38-.02-.54-.08-.16-.68-1.64-.93-2.24-.25-.6-.5-.5-.68-.51l-.58-.01c-.2 0-.52.07-.8.38-.27.31-1.05 1.03-1.05 2.5s1.08 2.9 1.24 3.11c.15.2 2.12 3.36 5.13 4.7.72.31 1.29.49 1.73.63.73.23 1.4.2 1.93.12.59-.09 1.77-.72 2.02-1.42.25-.7.25-1.28.18-1.41-.07-.13-.27-.2-.57-.36Z"/>
    </svg>`;
  btnGroup.appendChild(shareLink);
}


                const exampleBtn = card.querySelector('.example-btn');
                if (exampleBtn) {
                    exampleBtn.addEventListener('click', () => {
                       showExampleModal(item.name, item.syntax, item.example);
                    });
                }

                return card;
            };
            
            const renderContent = () => {
                contentGrid.innerHTML = '';
                const searchTerm = searchInput.value.toLowerCase();
                let data;

                if (currentTab === 'formulas') data = formulas;
                else if (currentTab === 'shortcuts') data = shortcuts;
                else data = [...formulas.filter(f => f.favorite), ...shortcuts.filter(s => s.favorite)];

                // --- More Powerful Search Logic with Synonyms ---
                const searchKeywords = searchTerm.split(' ').filter(k => k.length > 0);

                const filteredData = data.filter(item => {
                    const inCategory = activeCategory === 'All' || item.category === activeCategory || currentTab === 'favorites';
                    if (!inCategory) return false;

    		// ⬇️ NEW — Level filter (these are the 4 functional lines)
   		const inLevel =
        	activeLevel === 'All' ||
        	item.level === activeLevel;
    		if (!inLevel) return false;


                    const startsWithLetter = activeLetter === 'All' || item.name.toLowerCase().startsWith(activeLetter.toLowerCase());
                    if (!startsWithLetter && currentTab === 'formulas') return false;
                    
                    if (searchKeywords.length === 0) return true;

                    const searchableText = [
                        item.name,
                        item.category,
                        item.description,
                        item.syntax || item.keys || ''
                    ].join(' ').toLowerCase();

                    // Check if EVERY keyword (or one of its synonyms) is present
                    return searchKeywords.every(keyword => {
                        const synonyms = synonymMap[keyword] || [keyword];
                        return synonyms.some(synonym => searchableText.includes(synonym));
                    });
                });

		// ✅ Sort filtered data alphabetically by name
		filteredData.sort((a, b) => a.name.localeCompare(b.name));


                noResults.classList.toggle('hidden', filteredData.length > 0);
                
                filteredData.forEach(item => {
                    const type = formulas.includes(item) ? 'formula' : 'shortcut';
                    contentGrid.appendChild(createCard(item, type));
                });
            };
            
            const renderCategoryFilters = () => {
                categoryFiltersContainer.innerHTML = '';
                if (currentTab === 'formulas' || currentTab === 'shortcuts') {
                    const sourceData = currentTab === 'formulas' ? formulas : shortcuts;
                    const categories = ['All', ...new Set(sourceData.map(item => item.category).filter(Boolean))];
                    categories.forEach(category => {
                        const btn = document.createElement('button');
                        btn.textContent = category;
                        const isActive = activeCategory === category;
                        // Added 'category-btn' and 'active-category' for new 3D styling
                        btn.className = `category-btn px-3 py-1 text-sm font-medium rounded-full ${isActive ? 'active-category bg-blue-500 text-white' : 'bg-gray-200 dark:bg-gray-700'}`;
                        btn.onclick = () => {
                            activeCategory = category;
                            renderCategoryFilters();
                            renderContent();
                        };
                        categoryFiltersContainer.appendChild(btn);
                    });
                }
            };
            
            const renderAlphabetFilters = () => {
                alphabetFiltersContainer.innerHTML = '';
                if (currentTab === 'formulas') {
                    alphabetFiltersContainer.classList.remove('hidden');
                    const alphabet = ['All', ...'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.split('')];
                    alphabet.forEach(letter => {
                        const btn = document.createElement('button');
                        btn.textContent = letter;
                        const isActive = activeLetter === letter;
                        btn.className = `alphabet-btn w-8 h-8 sm:w-9 sm:h-9 text-xs sm:text-sm font-medium rounded-md ${isActive ? 'active-alphabet bg-blue-500 text-white' : 'bg-gray-200 dark:bg-gray-700'}`;
                        btn.onclick = () => {
                            activeLetter = letter;
                            renderAlphabetFilters(); // Re-render to show active state
                            renderContent();
                        };
                        alphabetFiltersContainer.appendChild(btn);
                    });
                } else {
                     alphabetFiltersContainer.classList.add('hidden');
                }
            };

            function toggleFavorite(e) {
                const btn = e.currentTarget;
                const name = btn.dataset.name;
                const type = btn.dataset.type;
                let item, favoriteList, storageKey;

                if (type === 'formula') {
                    item = formulas.find(f => f.name === name);
                    favoriteList = favoriteFormulas;
                    storageKey = 'favoriteFormulas';
                } else {
                    item = shortcuts.find(s => s.name === name);
                    favoriteList = favoriteShortcuts;
                    storageKey = 'favoriteShortcuts';
                }

                if (item) {
                    item.favorite = !item.favorite;
                    btn.classList.toggle('favorited', item.favorite);
                    
                    if (item.favorite) {
                        if (!favoriteList.includes(name)) favoriteList.push(name);
                    } else {
                        const index = favoriteList.indexOf(name);
                        if (index > -1) favoriteList.splice(index, 1);
                    }
                    localStorage.setItem(storageKey, JSON.stringify(favoriteList));
                    
                    if(currentTab === 'favorites') renderContent();
                }
            }

            function showExampleModal(name, syntax, example) {
                modalTitle.textContent = name;
                modalSyntax.textContent = syntax;
                modalExample.textContent = example;
                exampleModal.classList.remove('hidden');
                exampleModal.classList.add('flex');
            }

            function hideExampleModal() {
                exampleModal.classList.add('hidden');
                exampleModal.classList.remove('flex');
            }

            function copyToClipboard(text) {
                const textArea = document.createElement("textarea");
                textArea.value = text;
                document.body.appendChild(textArea);
                textArea.select();
                try {
                    document.execCommand('copy');
                    copiedToast.classList.add('show');
                    setTimeout(() => copiedToast.classList.remove('show'), 2000);
                } catch (err) {
                    console.error('Failed to copy text: ', err);
                }
                document.body.removeChild(textArea);
            }

            function updateTabSlider(activeTab) {
                if (activeTab) {
                    const left = activeTab.offsetLeft;
                    const width = activeTab.offsetWidth;
                    activeTabSlider.style.width = `${width}px`;
                    activeTabSlider.style.transform = `translateX(${left}px)`;
                }
            }

            searchInput.addEventListener('input', renderContent);

            tabs.forEach(tab => {
                tab.addEventListener('click', () => {
                    currentTab = tab.dataset.tab;
                    activeCategory = 'All';
                    activeLetter = 'All'; // Reset letter filter on tab change
                    tabs.forEach(t => t.classList.remove('active-tab'));
                    tab.classList.add('active-tab');
                    
                    updateTabSlider(tab);
                    renderAlphabetFilters();
                    renderCategoryFilters();
                    renderContent();
                });
            });

            // --- Modal Closers ---
            closeModalBtn.addEventListener('click', hideExampleModal);
            exampleModal.addEventListener('click', (e) => {
                if (e.target === exampleModal) { // only if clicking the overlay
                    hideExampleModal();
                }
            });

            // --- Theme Toggler ---
            const themeToggle = document.getElementById('theme-toggle');
            const lightIcon = document.getElementById('theme-icon-light');
            const darkIcon = document.getElementById('theme-icon-dark');
            const docElement = document.documentElement;

            const currentTheme = localStorage.getItem('theme') || 'dark';
            docElement.classList.toggle('dark', currentTheme === 'dark');
            lightIcon.classList.toggle('hidden', currentTheme === 'dark');
            darkIcon.classList.toggle('hidden', currentTheme === 'light');

            themeToggle.addEventListener('click', () => {
                const isDark = docElement.classList.toggle('dark');
                localStorage.setItem('theme', isDark ? 'dark' : 'light');
                lightIcon.classList.toggle('hidden');
                darkIcon.classList.toggle('hidden');
            });

            // Initial Render and UI setup
            const savedUser = JSON.parse(localStorage.getItem('excelVaultCurrentUser'));
	    if (savedUser) {
            updateUserUI(savedUser);
	    } else {
	    updateUserUI(null);
	    document.getElementById('greeting').textContent = `${timeOfDayGreeting}, Guest!`;
	    }
            renderAlphabetFilters();
            renderCategoryFilters();
            renderContent();
            // Set initial position of tab slider
            updateTabSlider(document.querySelector('.tab-btn.active-tab'));
        });
    </script>
<script>
// 1) Exact category list (first item is the "All" option shown on the button)
const CATEGORY_LIST = [
  'All Categories',
  'Core Excel Starter Kit',
  'Basics & Logic',
  'Array & Filter',
  'Error Handling',
  'Database & Information',
  'Advanced Lookup',
  'Conditional Math',
  'Dashboard & Sheet Info',
  'Dynamic Array & Lambda',
  'Basic Math & Logical Checks',
  'Date & Time',
  'Financial & Banking',
  'Loan & Credit',
  'Investment & Returns',
  'Statistical',
  'Text & Formatting',
  'Count & Frequency',
  'Math & Rounding',
  'Logical & Comparison',
  'Lookup & Reference',
  'Probability & Statistical Advanced',
  'Logical & Audit Helper',
  'Custom Financial Utilities',
  'Advanced Statistical & Analytical Tools',
  'Array Math & Matrix Operations',
  'Dynamic Data Cleaning & Transformation',
  'Business Calendar Automation',
  'Financial Modelling',
  'Power Query / Data Model Helpers',
  'Bank Branch Performance Metrics',
  'Data Validation & List Automation',
  'Inventory & Procurement Management',
  'Text Formatting & Number Display',
  'MSME Loan Monitoring',
  'Agriculture Credit Analysis',
  'Deposit Mobilization & Analysis',
  'Branch Ranking & Scorecard',
  'Operational Efficiency Metrics',
  'Profit & Loss Computation',
  'Expense & Budget Monitoring',
  'Banking & Loan Analysis',
  'Data Analysis Helper',
  'Engineering, Scientific & Math',
  'Power Query / Data Model',
  'Dashboard Formula Tricks',
  'Rarely Known & Handy',
  'Excel → PowerPoint',
  'Excel + AI Integration',
  'Excel 365 AI (Preview)',
  'Index & Search Helpers',
  'Text',
  'Financial',
  'Engineering & Complex',
  'Conversions & Bases',
  'Math & Trig',
  'Measurement & Unit Conversion',
  'Temperature & Physical Constants',
  'Household & Utility',
  'Decision Support',
  'Simulation',
  'Scenario Analysis',
  'Web & API',
  'AI & Copilot',
  'Matrix & Geometry',
  'Date Intelligence',
  'Data Science',
  'Banking KPI',
  'Logical',
  'Optimization',
  'Measurement',
  'Math',
  'Measurement & Geometry',
  'Math Constant',
  'Geometry',
  'HR & Operations',
  'HR & Payroll',
  'Array Math',
  'Text Analytics',
  'Market & Finance',
  'Finance',
  'Loan',
  'Data Analysis',
  'Forecasting',
  'Probability',
  'Operations & Supply Chain',
  'Web Integration',
  'Information',
  'Cross-Sheet & File',
  'Business Math',
  'Time & Shift',
  'Inventory',
  'Banking & Loan',
  'Risk & Portfolio',
  'Dynamic Filter',
  'Automation',
  'Audit',
  'Forecast',
  'Performance',
  'Conditional Formatting'
];

(function initCategoryDropdown() {
  const btn   = document.getElementById('categoryDropdownBtn');
  const list  = document.getElementById('categoryDropdownList');
  // Prevent internal scroll/press events from closing the dropdown
['pointerdown', 'wheel', 'touchstart', 'touchmove'].forEach(evt => {
  list.addEventListener(evt, (e) => {
    e.stopPropagation();   // stop this event from reaching the global listener
  }, { passive: false });
});

  const label = document.getElementById('selectedCategoryText');
  if (!btn || !list || !label) return; // safety

  // Build menu items
  list.innerHTML = CATEGORY_LIST.map(cat => `
    <button type="button" data-cat="${cat}"
            class="w-full text-left px-4 py-2 text-sm hover:bg-gray-100 dark:hover:bg-gray-700">
      ${cat}
    </button>
  `).join('');

  // Toggle open/close
  btn.addEventListener('click', (e) => {
    e.stopPropagation();
    list.classList.toggle('hidden');
  });

  // Click outside closes it
  document.addEventListener('pointerdown', (e) => {
  const t = e.target;
  // Close only if click/press is outside BOTH the button and the dropdown
  if (!btn.contains(t) && !list.contains(t)) {
    list.classList.add('hidden');
  }
});


  // Handle selection
  list.querySelectorAll('button[data-cat]').forEach(item => {
    item.addEventListener('click', () => {
      const chosen = item.getAttribute('data-cat');
      // Use the Step-1 hook if present; else fallback to label only
      if (typeof window.setCategory === 'function') {
        window.setCategory(chosen === 'All Categories' ? 'All' : chosen);
      } else {
        label.textContent = chosen; // visual only if hook missing
      }
      list.classList.add('hidden');
    });
  });
})();
</script>

</body>
</html>





