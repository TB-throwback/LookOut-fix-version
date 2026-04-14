plugins {
    id("com.android.application")
}

val pwaDir = rootProject.projectDir.parentFile.resolve("pwa")
val generatedPwaAssetsDir = layout.buildDirectory.dir("generated/pwaAssets").get().asFile

android {
    namespace = "io.github.tbthrowback.lookoutfixversion"
    compileSdk {
        version = release(36) {
            minorApiLevel = 1
        }
    }

    defaultConfig {
        applicationId = "io.github.tbthrowback.lookoutfixversion"
        minSdk = 24
        targetSdk = 36
        versionCode = 1
        versionName = "1.0"

        testInstrumentationRunner = "androidx.test.runner.AndroidJUnitRunner"
    }

    buildTypes {
        release {
            isMinifyEnabled = false
            proguardFiles(
                getDefaultProguardFile("proguard-android-optimize.txt"),
                "proguard-rules.pro"
            )
        }
    }
    compileOptions {
        sourceCompatibility = JavaVersion.VERSION_11
        targetCompatibility = JavaVersion.VERSION_11
    }

    sourceSets.getByName("main").assets.directories.add(generatedPwaAssetsDir.absolutePath)
}

tasks.register<Exec>("buildPwaAssets") {
    workingDir = pwaDir
    commandLine("node", "build.mjs")
    inputs.dir(pwaDir)
    outputs.dir(pwaDir.resolve("dist"))
}

tasks.register<Copy>("copyPwaAssets") {
    dependsOn("buildPwaAssets")
    from(pwaDir.resolve("dist"))
    into(generatedPwaAssetsDir)
}

tasks.named("preBuild") {
    dependsOn("copyPwaAssets")
}

dependencies {
    implementation("androidx.core:core-ktx:1.10.1")
    implementation("androidx.appcompat:appcompat:1.6.1")
    implementation("androidx.activity:activity:1.8.0")
    testImplementation("junit:junit:4.13.2")
    androidTestImplementation("androidx.test.ext:junit:1.1.5")
    androidTestImplementation("androidx.test.espresso:espresso-core:3.5.1")
}