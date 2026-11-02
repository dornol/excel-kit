plugins {
    `java-library`
    alias(libs.plugins.vanniktech.publish)
    id("signing")
    alias(libs.plugins.spring.dependency.management)
}

description = "Spring MVC helpers for excel-kit"

dependencyManagement {
    imports {
        mavenBom("org.springframework.boot:spring-boot-dependencies:${libs.versions.spring.boot.get()}")
    }
}

java {
    sourceCompatibility = JavaVersion.VERSION_17
    targetCompatibility = JavaVersion.VERSION_17
}

dependencies {
    api(project(":kit"))

    compileOnlyApi("org.springframework:spring-webmvc")
    compileOnlyApi(libs.poi.ooxml)
    compileOnlyApi(libs.jakarta.validation.api)
    compileOnlyApi(libs.jspecify)

    testImplementation(platform(libs.junit.bom))
    testImplementation(libs.junit.jupiter)
    testImplementation(libs.poi.ooxml)
    testImplementation(libs.opencsv)
    testImplementation(libs.slf4j.simple)
    testImplementation(libs.jakarta.validation.api)
    testImplementation("org.springframework:spring-webmvc")
    testImplementation("org.springframework:spring-test")
    testRuntimeOnly(libs.junit.platform.launcher)
}

tasks.withType<Test> {
    useJUnitPlatform()
}

tasks.withType<Javadoc>().configureEach {
    options.encoding = "UTF-8"
    (options as org.gradle.external.javadoc.StandardJavadocDocletOptions)
        .addStringOption("Xdoclint:all,-missing", "-quiet")
}

mavenPublishing {
    signAllPublications()
    publishToMavenCentral(automaticRelease = true)

    coordinates("io.github.dornol", "excel-kit-spring", project.version.toString())

    pom {
        // Maven has no compileOnlyApi scope. Preserve the optional integration
        // contract instead of forcing POI/Spring/Validation onto consumers.
        withXml {
            val optionalIntegrations = setOf(
                "org.springframework:spring-webmvc", "org.apache.poi:poi-ooxml",
                "jakarta.validation:jakarta.validation-api", "org.jspecify:jspecify"
            )
            val root = asElement()
            val dependencies = root.getElementsByTagName("dependency")
            for (i in 0 until dependencies.length) {
                val dependency = dependencies.item(i) as org.w3c.dom.Element
                if (dependency.parentNode.parentNode != root) continue
                val group = dependency.getElementsByTagName("groupId").item(0).textContent
                val artifact = dependency.getElementsByTagName("artifactId").item(0).textContent
                if ("$group:$artifact" in optionalIntegrations) {
                    dependency.appendChild(root.ownerDocument.createElement("optional").apply {
                        textContent = "true"
                    })
                }
            }
        }
        name = "excel-kit-spring"
        description = "Spring MVC helpers for excel-kit"
        url = "https://github.com/dornol/excel-kit/"

        licenses {
            license {
                name = "MIT"
                url = "https://github.com/dornol/excel-kit/blob/main/LICENSE"
            }
        }

        developers {
            developer {
                id = "dornol"
                name = "dhkim"
                email = "dhkim@dornol.dev"
                url = "https://github.com/dornol/"
            }
        }

        scm {
            url = "https://github.com/dornol/excel-kit/"
            connection = "scm:git:git://github.com/dornol/excel-kit.git"
            developerConnection = "scm:git:ssh://git@github.com/dornol/excel-kit.git"
        }
    }
}
