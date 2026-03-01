plugins {
    id("java")
    alias(libs.plugins.vanniktech.publish)
    id("signing")
}

java {
    sourceCompatibility = JavaVersion.VERSION_17
    targetCompatibility = JavaVersion.VERSION_17
}

dependencies {
    compileOnly(libs.poi.ooxml)
    compileOnly(libs.slf4j.api)
    compileOnly(libs.jakarta.validation.api)
    compileOnly(libs.opencsv)
    compileOnly(libs.jspecify)

    testImplementation(platform(libs.junit.bom))
    testImplementation(libs.junit.jupiter)
    testImplementation(libs.poi.ooxml)
    testImplementation(libs.opencsv)
    testImplementation(libs.slf4j.simple)
    testImplementation(libs.hibernate.validator)
    testRuntimeOnly(libs.junit.platform.launcher)
}

tasks.test {
    useJUnitPlatform()
}

tasks.withType<Javadoc>().configureEach {
    options.encoding = "UTF-8"
}

mavenPublishing {
    signAllPublications()
    publishToMavenCentral()

    coordinates("io.github.dornol", "excel-kit", project.version.toString())

    pom {
        name = "excel-kit"
        description = "Simple Excel download and upload kit"
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
