plugins {
    alias(libs.plugins.android.application)
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

    sourceSets.getByName("main").assets.setSrcDirs(listOf(generatedPwaAssetsDir))
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
    implementation(libs.androidx.core.ktx)
    implementation(libs.androidx.appcompat)
    implementation(libs.androidx.activity)
    testImplementation(libs.junit)
    androidTestImplementation(libs.androidx.junit)
    androidTestImplementation(libs.androidx.espresso.core)
}