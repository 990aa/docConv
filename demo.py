"""
Demo Python script to showcase syntax highlighting in DOCX conversion.
This file demonstrates various Python features with proper indentation.
"""
from typing import List, Dict, Optional
import asyncio

class DataProcessor:
    """A sample class to demonstrate code highlighting."""
    
    def __init__(self, name: str):
        self.name = name
        self.data: List[int] = []
    
    def add_data(self, value: int) -> None:
        """Add a value to the data list."""
        self.data.append(value)
        print(f"Added {value} to {self.name}")
    
    def calculate_statistics(self) -> Dict[str, float]:
        """Calculate basic statistics from the data."""
        if not self.data:
            return {"mean": 0.0, "sum": 0.0, "count": 0}
        
        total = sum(self.data)
        count = len(self.data)
        mean = total / count
        
        return {
            "mean": mean,
            "sum": total,
            "count": count,
            "min": min(self.data),
            "max": max(self.data)
        }


async def fetch_data(url: str) -> Optional[dict]:
    """Async function to fetch data."""
    try:
        # Simulated async operation
        await asyncio.sleep(0.1)
        return {"status": "success", "url": url}
    except Exception as e:
        print(f"Error fetching data: {e}")
        return None


def main():
    """Main entry point."""
    # Create processor instance
    processor = DataProcessor("SampleProcessor")
    
    # Add some data with list comprehension
    values = [x**2 for x in range(10) if x % 2 == 0]
    for val in values:
        processor.add_data(val)
    
    # Calculate and display statistics
    stats = processor.calculate_statistics()
    print("\nStatistics:")
    for key, value in stats.items():
        print(f"  {key}: {value}")
    
    # Lambda function example
    multiply = lambda x, y: x * y
    result = multiply(5, 3)
    print(f"\nLambda result: {result}")
    
    # Dictionary and string formatting
    config = {
        "host": "localhost",
        "port": 8080,
        "debug": True
    }
    print(f"\nServer config: {config}")


if __name__ == "__main__":
    main()
