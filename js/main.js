/**
 * الملف الرئيسي للنظام الشمسي التفاعلي ثلاثي الأبعاد
 */

// متغيرات عامة
let scene, camera, renderer, controls;
let planets = {};
let orbits = {};
let showOrbits = true;
let animationSpeed = 1;
let clock = new THREE.Clock();

// تهيئة المشهد
function init() {
    // إنشاء المشهد
    scene = new THREE.Scene();
    
    // إنشاء الكاميرا
    camera = new THREE.PerspectiveCamera(75, window.innerWidth / window.innerHeight, 0.1, 1000);
    camera.position.set(0, 30, 90);
    
    // إنشاء المُصيِّر
    renderer = new THREE.WebGLRenderer({ antialias: true });
    renderer.setSize(window.innerWidth, window.innerHeight);
    renderer.setPixelRatio(window.devicePixelRatio);
    document.getElementById('canvas-container').appendChild(renderer.domElement);
    
    // إضافة التحكم بالكاميرا
    controls = new THREE.OrbitControls(camera, renderer.domElement);
    controls.enableDamping = true;
    controls.dampingFactor = 0.05;
    controls.minDistance = 10;
    controls.maxDistance = 150;
    
    // إضافة الإضاءة
    const ambientLight = new THREE.AmbientLight(0x404040);
    scene.add(ambientLight);
    
    const sunLight = new THREE.PointLight(0xffffff, 2, 300);
    scene.add(sunLight);
    
    // إضافة خلفية النجوم
    createStarBackground();
    
    // إنشاء الكواكب
    createPlanets();
    
    // إضافة أحداث التفاعل
    addEventListeners();
    
    // بدء حلقة الرسم
    animate();
    
    // إخفاء شاشة التحميل بعد 3 ثوانٍ
    setTimeout(() => {
        document.getElementById('loading-screen').style.display = 'none';
    }, 3000);
}

// إنشاء خلفية النجوم
function createStarBackground() {
    const starsGeometry = new THREE.BufferGeometry();
    const starsMaterial = new THREE.PointsMaterial({
        color: 0xffffff,
        size: 0.7,
        sizeAttenuation: false
    });
    
    const starsVertices = [];
    for (let i = 0; i < 10000; i++) {
        const x = (Math.random() - 0.5) * 2000;
        const y = (Math.random() - 0.5) * 2000;
        const z = (Math.random() - 0.5) * 2000;
        starsVertices.push(x, y, z);
    }
    
    starsGeometry.setAttribute('position', new THREE.Float32BufferAttribute(starsVertices, 3));
    const starField = new THREE.Points(starsGeometry, starsMaterial);
    scene.add(starField);
}

// إنشاء الكواكب
function createPlanets() {
    const textureLoader = new THREE.TextureLoader();
    
    // إنشاء كل كوكب من بيانات الكواكب
    for (const [key, planet] of Object.entries(planetData)) {
        // تحميل القوام
        const texture = textureLoader.load(planet.texture, function() {
            // تحديث شريط التقدم أو إظهار رسالة عند اكتمال التحميل
        });
        
        // إنشاء الكوكب
        const geometry = new THREE.SphereGeometry(planet.radius, 32, 32);
        const material = new THREE.MeshPhongMaterial({
            map: texture,
            shininess: key === 'sun' ? 0 : 30
        });
        
        const mesh = new THREE.Mesh(geometry, material);
        
        // إضافة الكوكب إلى المشهد
        scene.add(mesh);
        
        // تخزين مرجع الكوكب
        planets[key] = {
            mesh: mesh,
            data: planet,
            angle: Math.random() * Math.PI * 2, // زاوية عشوائية للبدء
            rotationSpeed: 2 * Math.PI / (planet.rotationPeriod * 24) // سرعة الدوران حول نفسه
        };
        
        // إنشاء مدار الكوكب (إذا لم يكن الشمس)
        if (key !== 'sun') {
            const orbitGeometry = new THREE.RingGeometry(planet.distance - 0.1, planet.distance + 0.1, 128);
            const orbitMaterial = new THREE.MeshBasicMaterial({
                color: 0x444444,
                side: THREE.DoubleSide,
                transparent: true,
                opacity: 0.3
            });
            const orbit = new THREE.Mesh(orbitGeometry, orbitMaterial);
            orbit.rotation.x = Math.PI / 2;
            scene.add(orbit);
            orbits[key] = orbit;
        }
        
        // إضافة حلقات لكوكب زحل
        if (key === 'saturn') {
            const ringTexture = textureLoader.load(planet.ringTexture);
            const ringGeometry = new THREE.RingGeometry(planet.radius + 0.5, planet.radius * planet.ringSize, 64);
            const ringMaterial = new THREE.MeshBasicMaterial({
                map: ringTexture,
                side: THREE.DoubleSide,
                transparent: true,
                opacity: 0.8
            });
            const ring = new THREE.Mesh(ringGeometry, ringMaterial);
            ring.rotation.x = Math.PI / 2;
            planets[key].ring = ring;
            planets[key].mesh.add(ring);
        }
    }
}

// تحديث مواقع الكواكب
function updatePlanets(deltaTime) {
    for (const [key, planet] of Object.entries(planets)) {
        // تحديث موقع الكوكب
        if (key !== 'sun') {
            // حساب سرعة المدار
            const orbitalSpeed = 2 * Math.PI / (planet.data.orbitalPeriod * 24);
            
            // تحديث زاوية الكوكب
            planet.angle += orbitalSpeed * deltaTime * animationSpeed;
            
            // حساب الموقع الجديد
            const x = Math.cos(planet.angle) * planet.data.distance;
            const z = Math.sin(planet.angle) * planet.data.distance;
            
            planet.mesh.position.set(x, 0, z);
        } else {
            // الشمس في المركز
            planet.mesh.position.set(0, 0, 0);
        }
        
        // تدوير الكوكب حول نفسه
        planet.mesh.rotation.y += planet.rotationSpeed * deltaTime * animationSpeed;
    }
}

// حلقة الرسم الرئيسية
function animate() {
    requestAnimationFrame(animate);
    
    const deltaTime = clock.getDelta();
    
    // تحديث مواقع الكواكب
    updatePlanets(deltaTime);
    
    // تحديث التحكم بالكاميرا
    controls.update();
    
    // رسم المشهد
    renderer.render(scene, camera);
}

// إضافة أحداث التفاعل
function addEventListeners() {
    // تغيير حجم النافذة
    window.addEventListener('resize', () => {
        camera.aspect = window.innerWidth / window.innerHeight;
        camera.updateProjectionMatrix();
        renderer.setSize(window.innerWidth, window.innerHeight);
    });
    
    // إظهار/إخفاء المدارات
    document.getElementById('toggle-orbits').addEventListener('click', () => {
        showOrbits = !showOrbits;
        for (const orbit of Object.values(orbits)) {
            orbit.visible = showOrbits;
        }
    });
    
    // تسريع/إبطاء الحركة
    document.getElementById('toggle-speed').addEventListener('click', () => {
        animationSpeed = (animationSpeed === 1) ? 5 : 1;
    });
    
    // إعادة ضبط المنظور
    document.getElementById('reset-view').addEventListener('click', () => {
        camera.position.set(0, 30, 90);
        controls.target.set(0, 0, 0);
    });
    
    // إظهار معلومات الكواكب
    document.getElementById('toggle-info').addEventListener('click', () => {
        const infoPanel = document.getElementById('info-panel');
        infoPanel.classList.toggle('hidden');
        
        // عرض معلومات الشمس افتراضياً
        if (!infoPanel.classList.contains('hidden')) {
            showPlanetInfo('sun');
        }
    });
    
    // إغلاق لوحة المعلومات
    document.querySelector('.close-btn').addEventListener('click', () => {
        document.getElementById('info-panel').classList.add('hidden');
    });
    
    // اختيار الكواكب
    const planetOptions = document.querySelectorAll('.planet-option');
    planetOptions.forEach(option => {
        option.addEventListener('click', () => {
            const planetKey = option.getAttribute('data-planet');
            
            // عرض معلومات الكوكب
            showPlanetInfo(planetKey);
            
            // تحريك الكاميرا إلى الكوكب
            focusOnPlanet(planetKey);
            
            // تحديث الكوكب النشط
            planetOptions.forEach(opt => opt.classList.remove('active'));
            option.classList.add('active');
        });
    });
}

// عرض معلومات الكوكب
function showPlanetInfo(planetKey) {
    const planet = planetData[planetKey];
    document.getElementById('planet-name').textContent = planet.name;
    document.getElementById('planet-info').innerHTML = planet.info;
    document.getElementById('info-panel').classList.remove('hidden');
}

// التركيز على كوكب معين
function focusOnPlanet(planetKey) {
    const planet = planets[planetKey];
    
    // تحريك الكاميرا إلى موقع الكوكب
    const targetPosition = new THREE.Vector3().copy(planet.mesh.position);
    
    // تعديل المسافة حسب حجم الكوكب
    const distance = planet.data.radius * 10 + 5;
    
    // حساب موقع الكاميرا الجديد
    const cameraPosition = new THREE.Vector3(
        targetPosition.x + distance,
        distance / 2,
        targetPosition.z + distance
    );
    
    // تحريك الكاميرا بسلاسة
    new TWEEN.Tween(camera.position)
        .to(cameraPosition, 1000)
        .easing(TWEEN.Easing.Quadratic.Out)
        .start();
    
    // تحديث هدف التحكم
    controls.target.copy(targetPosition);
}

// بدء التطبيق
window.addEventListener('load', init);
