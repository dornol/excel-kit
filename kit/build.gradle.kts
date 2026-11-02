plugins {
    id("java")
    id("jacoco")
    alias(libs.plugins.vanniktech.publish)
    id("signing")
}

java {
    sourceCompatibility = JavaVersion.VERSION_17
    targetCompatibility = JavaVersion.VERSION_17
}

dependencies {
    val poiVersion = providers.gradleProperty("poiVersion").orNull
    val opencsvVersion = providers.gradleProperty("opencsvVersion").orNull

    compileOnly(if (poiVersion != null) "org.apache.poi:poi-ooxml:$poiVersion" else libs.poi.ooxml)
    compileOnly(libs.slf4j.api)
    compileOnly(libs.jakarta.validation.api)
    compileOnly(if (opencsvVersion != null) "com.opencsv:opencsv:$opencsvVersion" else libs.opencsv)
    compileOnly(libs.jspecify)

    testImplementation(platform(libs.junit.bom))
    testImplementation(libs.junit.jupiter)
    testImplementation(if (poiVersion != null) "org.apache.poi:poi-ooxml:$poiVersion" else libs.poi.ooxml)
    testImplementation(if (opencsvVersion != null) "com.opencsv:opencsv:$opencsvVersion" else libs.opencsv)
    testImplementation(libs.slf4j.simple)
    testImplementation(libs.jakarta.validation.api)
    testImplementation(libs.hibernate.validator)
    testRuntimeOnly(libs.junit.platform.launcher)
}

tasks.test {
    useJUnitPlatform {
        excludeTags("benchmark")
    }
    finalizedBy(tasks.jacocoTestReport)
}

// Exercise the published artifact with each supported POI runtime, without
// recompiling the library for that runtime's callback interface.
val poiRuntimeTests = listOf("5.2.5", "5.5.1", "6.0.0").map { poiVersion ->
    val suffix = poiVersion.replace(".", "_")
    val runtime = configurations.create("poiRuntime$suffix") {
        isCanBeConsumed = false
        isCanBeResolved = true
        extendsFrom(configurations.testRuntimeClasspath.get())
        resolutionStrategy.eachDependency {
            if (requested.group == "org.apache.poi") {
                useVersion(poiVersion)
            }
        }
    }
    tasks.register<Test>("testPoi$suffix") {
        group = "verification"
        description = "Test the compiled library JAR with POI $poiVersion"
        dependsOn(tasks.jar, tasks.testClasses)
        testClassesDirs = sourceSets.test.get().output.classesDirs
        classpath = files(sourceSets.test.get().output, tasks.jar.flatMap { it.archiveFile }, runtime)
        useJUnitPlatform {
            excludeTags("benchmark")
        }
    }
}

tasks.register("poiCompatibilityTest") {
    group = "verification"
    description = "Test the same library JAR with all supported POI runtimes"
    dependsOn(poiRuntimeTests)
}

tasks.register<Test>("benchmark") {
    group = "verification"
    description = "Run performance benchmarks"
    testClassesDirs = sourceSets.test.get().output.classesDirs
    classpath = sourceSets.test.get().runtimeClasspath
    useJUnitPlatform {
        includeTags("benchmark")
    }
    maxHeapSize = "1g"
    testLogging {
        showStandardStreams = true
    }
}

tasks.jacocoTestReport {
    dependsOn(tasks.test)
    reports {
        csv.required = true
        xml.required = true
    }
}

tasks.jacocoTestCoverageVerification {
    dependsOn(tasks.jacocoTestReport)
    violationRules {
        rule {
            limit {
                minimum = "0.70".toBigDecimal()
            }
        }
    }
}

tasks.check {
    dependsOn(tasks.jacocoTestCoverageVerification)
}

tasks.withType<Javadoc>().configureEach {
    options.encoding = "UTF-8"
    (options as org.gradle.external.javadoc.StandardJavadocDocletOptions)
        .addStringOption("Xdoclint:all,-missing", "-quiet")
}

mavenPublishing {
    signAllPublications()
    publishToMavenCentral(automaticRelease = true)

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
