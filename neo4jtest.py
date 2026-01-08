from neo4j import GraphDatabase

def test_connection(uri, user="neo4j", password="neo4j", encrypted=True):
    print(f"\n=== Testing {uri} (encrypted={encrypted}) ===")
    try:
        driver = GraphDatabase.driver(
            uri,
            auth=(user, password),
            encrypted=encrypted,
            connection_timeout=5  # fail fast, 5 seconds
        )
        with driver.session() as session:
            result = session.run("RETURN 'Connection successful' AS msg")
            print("OK:", result.single()["msg"])
    except Exception as e:
        print("ERROR:", repr(e))


endpoints = [
    # public URLS
    ("bolt://neo4j.genai.accentureanalytics.com:7687", True),
    ("bolt://neo4j.genai.accentureanalytics.com:7687", False),
    ("neo4j://neo4j.genai.accentureanalytics.com:7687", True),

    ("bolt://neo4j.genai.accentureanalytics.com", True),
    ("bolt://neo4j.genai.accentureanalytics.com", False),
    ("neo4j://neo4j.genai.accentureanalytics.com", True),

    # Private urls
    ("bolt://neo4j-container.bluesky-12f77b3a.eastus.azurecontainerapps.io:7687", True),
    ("bolt://neo4j-container.bluesky-12f77b3a.eastus.azurecontainerapps.io:7687", False),
]

for uri, enc in endpoints:
    test_connection(uri, encrypted=enc)
