module ToSpreadsheet
  class Context
    # Associating Axlsx entities and Nokogiri nodes
    module Pairing
      def assoc!(entity, node)
        @entity_to_node         ||= {}
        @node_to_entity         ||= {}
        @entity_to_node[entity] = node
        @node_to_entity[node]   = entity
      end

      def entity_from_node(node)
        @node_to_entity[node]
      end

      def node_from_entity(entity)
        @entity_to_node[entity]
      end

      def clear_assoc!
        @entity_to_node = {}
        @node_to_entity = {}
      end
    end
  end
end