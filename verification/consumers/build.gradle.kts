val kitVersion = providers.gradleProperty("kitVersion").get()
val repository = providers.gradleProperty("consumerRepository").get()
val pomOnly = providers.gradleProperty("pomOnly").map(String::toBoolean).orElse(false)

subprojects {
    apply(plugin = "application")
    repositories {
        maven {
            url = uri(repository)
            content { includeGroup("io.github.dornol") }
            if (pomOnly.get()) {
                metadataSources {
                    mavenPom()
                    ignoreGradleMetadataRedirection()
                }
            }
        }
        mavenCentral { content { excludeGroup("io.github.dornol") } }
    }
    extensions.configure<JavaPluginExtension> {
        toolchain.languageVersion.set(JavaLanguageVersion.of(JavaVersion.current().majorVersion))
    }
    tasks.withType<JavaCompile>().configureEach { options.release.set(17) }
    extensions.configure<JavaApplication> { mainClass.set("Consumer") }
    dependencies {
        add("implementation", "io.github.dornol:${if (name == "spring") "excel-kit-spring" else "excel-kit"}:$kitVersion")
        add("implementation", "org.slf4j:slf4j-api:2.0.20")
        when (name) {
            "excel", "spring" -> add("implementation", "org.apache.poi:poi-ooxml:6.0.0")
            "csv", "validation" -> add("implementation", "com.opencsv:opencsv:5.12.0")
        }
        if (name == "validation") {
            add("implementation", "org.hibernate.validator:hibernate-validator:9.1.5.Final")
        }
        if (name == "spring") {
            add("implementation", "org.springframework:spring-webmvc:7.0.10")
        }
    }
}

tasks.register("verifyConsumers") {
    dependsOn(subprojects.map { "${it.path}:run" })
}
