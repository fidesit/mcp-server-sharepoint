import com.vanniktech.maven.publish.JavadocJar
import com.vanniktech.maven.publish.KotlinJvm
import com.vanniktech.maven.publish.SonatypeHost

plugins {
    kotlin("jvm") version "2.1.20"
    id("com.gradleup.shadow") version "9.0.0-beta12"
    id("com.vanniktech.maven.publish") version "0.30.0"
}

group = "com.utisha"
version = "0.1.0"

java {
    toolchain {
        languageVersion = JavaLanguageVersion.of(21)
    }
}

repositories {
    mavenCentral()
}

dependencies {
    // MCP SDK — server framework for Model Context Protocol
    implementation("io.modelcontextprotocol.sdk:mcp:0.17.2")

    // Microsoft Graph SDK — SharePoint operations via Graph API
    implementation("com.microsoft.graph:microsoft-graph:6.29.0")

    // Azure Identity — MSAL4J-based authentication (client_credentials flow)
    implementation("com.azure:azure-identity:1.15.4")

    // Jackson JSR-310 — OffsetDateTime serialization in tool responses
    implementation("com.fasterxml.jackson.datatype:jackson-datatype-jsr310:2.18.3")

    // Logging
    implementation("org.slf4j:slf4j-api:2.0.16")
    runtimeOnly("ch.qos.logback:logback-classic:1.5.16")

    // Test
    testImplementation(kotlin("test"))
    testImplementation("io.mockk:mockk:1.13.13")
}

tasks.withType<org.jetbrains.kotlin.gradle.tasks.KotlinCompile> {
    compilerOptions {
        freeCompilerArgs.add("-Xjsr305=strict")
        jvmTarget.set(org.jetbrains.kotlin.gradle.dsl.JvmTarget.JVM_21)
    }
}

tasks.withType<Test> {
    useJUnitPlatform()
    maxParallelForks = (Runtime.getRuntime().availableProcessors() / 2).coerceAtLeast(1)
}

// Exclude integration tests from the default 'test' task
tasks.named<Test>("test") {
    useJUnitPlatform {
        excludeTags("integration")
    }
}

// Separate task for integration tests (requires SharePoint credentials)
tasks.register<Test>("integrationTest") {
    description = "Runs integration tests against a real SharePoint tenant."
    group = "verification"
    useJUnitPlatform {
        includeTags("integration")
    }
    testClassesDirs = sourceSets["test"].output.classesDirs
    classpath = sourceSets["test"].runtimeClasspath
    shouldRunAfter(tasks.named("test"))
}

// ── Fat JAR (for direct use: java -jar) ─────────────────────────────────

tasks.shadowJar {
    archiveBaseName.set("mcp-server-sharepoint")
    archiveClassifier.set("")
    archiveVersion.set(project.version.toString())
    mergeServiceFiles()
    manifest {
        attributes("Main-Class" to "com.utisha.mcp.sharepoint.MainKt")
    }
}

// Make 'build' produce the shadow JAR
tasks.named("build") {
    dependsOn(tasks.shadowJar)
}

// ── Maven Central Publishing ─────────────────────────────────────────────

mavenPublishing {
    configure(KotlinJvm(javadocJar = JavadocJar.Empty()))

    publishToMavenCentral(SonatypeHost.CENTRAL_PORTAL, automaticRelease = true)
    signAllPublications()

    coordinates(group.toString(), "mcp-server-sharepoint", version.toString())

    pom {
        name.set("mcp-server-sharepoint")
        description.set("A Model Context Protocol (MCP) server for Microsoft SharePoint. Browse sites, manage documents, and search files through AI assistants.")
        inceptionYear.set("2026")
        url.set("https://github.com/fidesit/mcp-server-sharepoint")

        licenses {
            license {
                name.set("The Apache License, Version 2.0")
                url.set("https://www.apache.org/licenses/LICENSE-2.0.txt")
                distribution.set("repo")
            }
        }

        developers {
            developer {
                id.set("fidesit")
                name.set("Fides IT")
                url.set("https://github.com/fidesit")
            }
        }

        scm {
            url.set("https://github.com/fidesit/mcp-server-sharepoint")
            connection.set("scm:git:git://github.com/fidesit/mcp-server-sharepoint.git")
            developerConnection.set("scm:git:ssh://git@github.com/fidesit/mcp-server-sharepoint.git")
        }
    }
}
