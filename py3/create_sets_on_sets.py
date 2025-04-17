from Grasshopper import DataTree
from Grasshopper.Kernel.Data import GH_Path
from collections import defaultdict

# Function to convert a Grasshopper DataTree to a list of (path, items) tuples
def tree_to_list_with_paths(input_tree):
    """Converts a Grasshopper DataTree into a list of (path, items) tuples."""
    all_branches = []
    for i in range(input_tree.BranchCount):
        path = input_tree.Path(i)
        branch = input_tree.Branch(i)
        all_branches.append((path, list(branch)))
    return all_branches

# Function to generate a hashable key for any item
def get_item_key(item):
    if isinstance(item, float):
        # Round floats to mitigate precision errors
        return ('float', round(item, 6))  # Adjust the number of decimals as needed
    elif isinstance(item, int):
        return ('int', item)
    elif isinstance(item, str):
        return ('str', item)
    elif isinstance(item, (tuple, list)):
        # Recursively generate keys for elements in sequences
        return (type(item).__name__, tuple(get_item_key(subitem) for subitem in item))
    else:
        # For other types, attempt to use the item directly if hashable
        try:
            hash(item)
            return (type(item).__name__, item)
        except TypeError:
            # As a last resort, use the string representation
            return (type(item).__name__, str(item))

# Function to compare the structures of two DataTrees
def compare_tree_structures(tree1, tree2):
    """Checks if two DataTrees have the same structure."""
    if tree1.BranchCount != tree2.BranchCount:
        return False
    for i in range(tree1.BranchCount):
        path1 = tree1.Path(i)
        path2 = tree2.Path(i)
        if path1 != path2:
            return False
        list1 = tree1.Branch(path1)
        list2 = tree2.Branch(path2)
        if len(list1) != len(list2):
            return False
    return True

# Function to group similar items within each branch, preserving paths
def group_similar_objects_with_paths(primary_branches, secondary_branches=None):
    """Groups identical items within primary_branches and optionally applies the same grouping to secondary_branches."""
    grouped_primary_branches = []
    grouped_secondary_branches = [] if secondary_branches is not None else None

    for idx, (path, primary_items) in enumerate(primary_branches):
        item_counts = defaultdict(list)
        for i, primary_item in enumerate(primary_items):
            item_key = get_item_key(primary_item)
            item_counts[item_key].append(primary_item)

        grouped_primary_list = list(item_counts.values())
        grouped_primary_branches.append((path, grouped_primary_list))

        # If secondary data is provided, group it synchronously
        if secondary_branches is not None:
            secondary_path, secondary_items = secondary_branches[idx]
            # Ensure paths match
            if path != secondary_path:
                raise ValueError("Error: Paths of primary and secondary data do not match.")
            secondary_counts = defaultdict(list)
            for i, secondary_item in enumerate(secondary_items):
                primary_item = primary_items[i]
                item_key = get_item_key(primary_item)
                secondary_counts[item_key].append(secondary_item)
            grouped_secondary_list = [secondary_counts[key] for key in item_counts.keys()]
            grouped_secondary_branches.append((path, grouped_secondary_list))

    return grouped_primary_branches, grouped_secondary_branches

# Function to convert a list of (path, grouped_list) back to a DataTree
def list_to_tree_with_paths(grouped_branches):
    output_tree = DataTree[object]()
    for path, grouped_list in grouped_branches:
        path_indices = [path[i] for i in range(path.Length)]
        for j, group in enumerate(grouped_list):
            new_path_indices = path_indices + [j]
            new_path = GH_Path(*new_path_indices)
            output_tree.AddRange(group, new_path)
    return output_tree

# Main script execution
# Ensure that 'input_data_tree' is connected as a DataTree input
# 'secondary_data_tree' is optional

# Step 1: Convert input DataTree to list of (path, items)
primary_branches = tree_to_list_with_paths(input_data_tree)

# Check if secondary_data_tree is provided and has data
if 'secondary_data_tree' in globals() and secondary_data_tree is not None and secondary_data_tree.BranchCount > 0:
    # Convert secondary DataTree to list of (path, items)
    secondary_branches = tree_to_list_with_paths(secondary_data_tree)
    # Step 2: Check if both DataTrees have the same structure
    if not compare_tree_structures(input_data_tree, secondary_data_tree):
        raise ValueError("Error: The input DataTrees have different structures. Both inputs must have the same tree structure for synchronous grouping.")
    # Step 3: Group items and apply synchronous grouping to secondary items
    grouped_primary_branches, grouped_secondary_branches = group_similar_objects_with_paths(
        primary_branches, secondary_branches
    )
    # Step 4: Convert grouped branches back to DataTrees
    output_data_tree = list_to_tree_with_paths(grouped_primary_branches)
    output = output_data_tree

    output_secondary_data_tree = list_to_tree_with_paths(grouped_secondary_branches)
    output_secondary = output_secondary_data_tree
else:
    # No secondary data provided; group only primary data
    grouped_primary_branches, _ = group_similar_objects_with_paths(primary_branches)
    # Convert grouped branches back to DataTree
    output_data_tree = list_to_tree_with_paths(grouped_primary_branches)
    output = output_data_tree
    # Ensure output_secondary is set to None
    output_secondary = None