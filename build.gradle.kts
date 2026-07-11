plugins {
    alias(libs.plugins.vanniktech.publish) apply false
    alias(libs.plugins.spring.dependency.management) apply false
}

allprojects {

    group = "io.github.dornol"
    version = "0.20.0"

    repositories {
        mavenCentral()
    }

    plugins.withType<JavaPlugin> {
        extensions.configure<JavaPluginExtension> {
            toolchain {
                languageVersion = JavaLanguageVersion.of(21)
            }
        }
    }
}
