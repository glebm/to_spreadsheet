module ToSpreadsheet
  class Context
    # Axlsx classes <-> Nokogiri table nodes round-tripping
    module Pairing
      def assoc!(entity, node)
        @entity_to_node         ||= {}
        @node_to_entity         ||= {}
        @entity_to_node[entity] = node
        @node_to_entity[node]   = entity
      end

      def to_xls_entity(node)
        @node_to_entity[node]
      end

      def to_xml_node(entity)
        @entity_to_node[entity]
      end

      def xml_node_and_xls_entity(entity)
        [@entity_to_node[entity], entity, @node_to_entity[entity]].compact
      end

      def clear_assoc!
        @entity_to_node = {}
        @node_to_entity = {}
      end
    end
  end
end