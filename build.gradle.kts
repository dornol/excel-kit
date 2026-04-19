plugins {
}

allprojects {

    group = "io.github.dornol"
    version = "0.16.15"

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