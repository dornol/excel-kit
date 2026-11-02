plugins {
    alias(libs.plugins.vanniktech.publish) apply false
    alias(libs.plugins.spring.dependency.management) apply false
}

allprojects {

    group = "io.github.dornol"
    version = "0.23.0"

    repositories {
        mavenCentral()
    }

    plugins.withType<JavaPlugin> {
        extensions.configure<JavaPluginExtension> {
            toolchain {
                languageVersion = JavaLanguageVersion.of(21)
            }
        }
        tasks.withType<JavaCompile>().configureEach {
            options.release.set(17)
        }
        val testVersion = providers.gradleProperty("testJavaVersion")
            .map(String::toInt).orElse(JavaVersion.current().majorVersion.toInt())
        val toolchains = extensions.getByType<JavaToolchainService>()
        tasks.withType<Test>().configureEach {
            javaLauncher.set(toolchains.launcherFor {
                languageVersion.set(testVersion.map(JavaLanguageVersion::of))
            })
            doFirst {
                logger.lifecycle("$path test JVM: ${javaLauncher.get().metadata.languageVersion} " +
                    "(${javaLauncher.get().executablePath})")
            }
        }
    }

    plugins.withId("maven-publish") {
        extensions.configure<PublishingExtension> {
            repositories.maven {
                name = "consumerTest"
                url = rootProject.layout.buildDirectory.dir("consumer-repository").get().asFile.toURI()
            }
        }
    }
}

tasks.register("printVersion") {
    val releaseVersion = project.version.toString()
    doLast { println(releaseVersion) }
}

tasks.register("verifyReleaseVersion") {
    group = "verification"
    description = "Verify the release tag matches all published module versions"
    val tag = providers.gradleProperty("releaseTag").orElse(providers.environmentVariable("GITHUB_REF_NAME"))
    val releaseVersion = project.version.toString()
    val publishedVersions = listOf(":kit", ":excel-kit-spring")
        .associateWith { project(it).version.toString() }
    doLast {
        val expected = "v$releaseVersion"
        check(tag.orNull == expected && expected.matches(Regex("v\\d+\\.\\d+\\.\\d+"))) {
            "Release tag ${tag.orNull} must match project version $expected"
        }
        publishedVersions.forEach { (module, version) ->
            check(version == releaseVersion) { "$module version differs from the release version" }
        }
    }
}
